<#
.SYNOPSIS
    Bulk creates and provisions Microsoft Teams Shared Space accounts using the official
    Microsoft 365 Admin Center user import template, enabling international calling.

.DESCRIPTION
    1. Parses official M365 Admin Center CSV header schema.
    2. Provisions Entra ID (Azure AD) user accounts using Microsoft Graph SDK.
    3. Dynamically identifies and assigns the Teams Shared Devices license SKU.
    4. Provisions enterprise voice and assigns E.164 phone numbers via Teams PowerShell.
    5. Grants policies required for International PSTN dialing:
       - Teams Calling Policy (AllowInternationalCalling = $true)
       - Teams Voice Routing Policy (PSTN usage records allowing international routes)
       - Teams Dial Plan (Normalization rules for international prefixes/exit codes)
       - Teams IP Phone Policy (Shared/Common Area Phone sign-in behavior)

.PARAMETER CsvPath
    The path to the Microsoft 365 Admin Center CSV file.

.PARAMETER CallingPolicy
    Name of the Teams Calling Policy configured with international dialing permissions.
    Default: 'AllowInternationalCalling'

.PARAMETER VoiceRoutingPolicy
    Name of the Teams Voice Routing Policy containing international PSTN usage records.
    (Applicable for Direct Routing / Operator Connect). Default: 'InternationalVoiceRouting'

.PARAMETER DialPlan
    Name of the tenant-level Dial Plan supporting international prefix normalization.
    Default: 'Global'

.PARAMETER IPPhonePolicy
    Name of the Teams IP Phone Policy (Common Area Phone profile).
    Default: 'CommonAreaPhone'

.PARAMETER PhoneNumberType
    PSTN connectivity model. Options: 'DirectRouting', 'CallingPlan', 'OperatorConnect'.
    Default: 'DirectRouting'

.NOTES
    Author:      Soutik Bera
    GitHub:      https://github.com/BeraSoutik7
    Repository:  PowerShell-M365-Starters

.EXAMPLE
    .\New-TeamsSharedDevicesBulk.ps1 -CsvPath ".\M365_Bulk_Users.csv" -CallingPolicy "International-Allowed" -VoiceRoutingPolicy "Global-Intl-Route"
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CsvPath,

    [Parameter(Mandatory = $false)]
    [string]$CallingPolicy = "AllowInternationalCalling",

    [Parameter(Mandatory = $false)]
    [string]$VoiceRoutingPolicy = "InternationalVoiceRouting",

    [Parameter(Mandatory = $false)]
    [string]$DialPlan = "Global",

    [Parameter(Mandatory = $false)]
    [string]$IPPhonePolicy = "CommonAreaPhone",

    [Parameter(Mandatory = $false)]
    [ValidateSet("DirectRouting", "CallingPlan", "OperatorConnect")]
    [string]$PhoneNumberType = "DirectRouting"
)

# -------------------------------------------------------------------------
# 1. Module & Session Validation
# -------------------------------------------------------------------------
$requiredModules = @("Microsoft.Graph.Users", "Microsoft.Graph.Identity.DirectoryManagement", "MicrosoftTeams")
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Throw "Missing prerequisite module: '$module'. Install via: Install-Module $module -Scope CurrentUser"
    }
}

# Graph Connection Verification
$mgContext = Get-MgContext -ErrorAction SilentlyContinue
if (-not $mgContext) {
    Write-Host "[AUTH] Connecting to Microsoft Graph SDK..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All" -NoWelcome
}

# Teams PowerShell Connection Verification
try {
    $null = Get-CsTenant -ErrorAction Stop
} catch {
    Write-Host "[AUTH] Connecting to Microsoft Teams PowerShell Module..." -ForegroundColor Cyan
    Connect-MicrosoftTeams | Out-Null
}

# -------------------------------------------------------------------------
# 2. License SKU Resolution (Teams Shared Devices)
# -------------------------------------------------------------------------
Write-Host "[CONFIG] Locating Teams Shared Devices license SKU..." -ForegroundColor Cyan
$sharedDeviceSku = Get-MgSubscribedSku | Where-Object {
    $_.SkuPartNumber -match "^(MCOCAP|TEAMS_SHARED_DEVICES)$" -and ($_.PrepaidUnits.Enabled - $_.ConsumedUnits) -gt 0
} | Select-Object -First 1

if (-not $sharedDeviceSku) {
    Throw "No available 'Teams Shared Devices' or 'MCOCAP' license SKU with unassigned seats found in this tenant."
}
Write-Host "[CONFIG] Using SKU: $($sharedDeviceSku.SkuPartNumber) (Available: $($sharedDeviceSku.PrepaidUnits.Enabled - $sharedDeviceSku.ConsumedUnits))" -ForegroundColor Green

# -------------------------------------------------------------------------
# 3. CSV Import & Data Normalization
# -------------------------------------------------------------------------
Write-Host "[INPUT] Importing CSV: $CsvPath" -ForegroundColor Cyan
$rawRows = Import-Csv -Path $CsvPath

if (-not $rawRows) {
    Throw "The supplied CSV file is empty or formatted incorrectly."
}

$successCount = 0
$failCount = 0

foreach ($row in $rawRows) {
    # Map official M365 Admin Center bulk user headers
    $displayName   = $row.'Name [displayName]'
    $upn           = $row.'User name [userPrincipalName]'
    $rawPassword   = $row.'Password'
    $forceChange   = if ($row.'Force change password' -match 'TRUE|true|1') { $true } else { $false }
    $givenName     = $row.'First name [givenName]'
    $surname       = $row.'Last name [surname]'
    $jobTitle      = $row.'Job title [jobTitle]'
    $department    = $row.'Department [department]'
    $usageLocation = $row.'Usage location [usageLocation]'
    $phone         = $row.'Business phone [telephoneNumber]'
    
    # Validation checks
    if ([string]::IsNullOrWhiteSpace($upn) -or [string]::IsNullOrWhiteSpace($displayName)) {
        Write-Warning "Skipping invalid row: Missing UPN or DisplayName."
        $failCount++
        continue
    }

    if ([string]::IsNullOrWhiteSpace($usageLocation)) {
        Write-Warning "[$upn] Usage location is required for license assignment. Defaulting to 'US'."
        $usageLocation = "US"
    }

    Write-Host "`n>>> Processing Shared Account: $displayName ($upn)" -ForegroundColor Yellow

    try {
        # -----------------------------------------------------------------
        # Step A: Check / Provision User in Entra ID
        # -----------------------------------------------------------------
        $existingUser = Get-MgUser -UserId$upn -ErrorAction SilentlyContinue

        if (-not $existingUser) {
            Write-Host "  [-] Creating Entra ID user object..." -ForegroundColor Gray
            
            $passwordProfile = @{
                Password                      = $rawPassword
                ForceChangePasswordNextSignIn = $forceChange
            }

            # Generate mail nickname from UPN prefix
            $nickname = ($upn -split "@")[0] -replace '[^a-zA-Z0-9]', ''

            $newUserParams = @{
                DisplayName       = $displayName
                UserPrincipalName = $upn
                MailNickname      = $nickname
                AccountEnabled    = $true
                PasswordProfile   = $passwordProfile
                UsageLocation     = $usageLocation
                GivenName         = $givenName
                Surname           = $surname
                JobTitle          = $jobTitle
                Department        = $department
            }

            if ($PSCmdlet.ShouldProcess($upn, "Create User")) {
                $userObj = New-MgUser @newUserParams
                Write-Host "  [+] User created successfully. ID: $($userObj.Id)" -ForegroundColor Green
            }
        } else {
            Write-Host "  [*] User already exists. Updating usage location if missing..." -ForegroundColor Gray
            $userObj =$existingUser
            if (-not $userObj.UsageLocation) {
                Update-MgUser -UserId $userObj.Id -UsageLocation$usageLocation
            }
        }

        # -----------------------------------------------------------------
        # Step B: Assign Teams Shared Devices License
        # -----------------------------------------------------------------
        $currentLicenses = (Get-MgUserLicenseDetail -UserId$userObj.Id).SkuPartNumber
        if ($currentLicenses -notcontains$sharedDeviceSku.SkuPartNumber) {
            Write-Host "  [-] Assigning Teams Shared Devices license..." -ForegroundColor Gray
            if ($PSCmdlet.ShouldProcess($userObj.UserPrincipalName, "Assign License $($sharedDeviceSku.SkuPartNumber)")) {
                Set-MgUserLicense -UserId $userObj.Id `
                                  -AddLicenses @(@{ SkuId = $sharedDeviceSku.SkuId }) `
                                  -RemoveLicenses @() | Out-Null
                Write-Host "  [+] License assigned: $($sharedDeviceSku.SkuPartNumber)" -ForegroundColor Green
            }
        } else {
            Write-Host "  [*] License already present: $($sharedDeviceSku.SkuPartNumber)" -ForegroundColor Gray
        }

        # Propagation delay for Teams provisioning backend
        Start-Sleep -Seconds 5

        # -----------------------------------------------------------------
        # Step C: Assign Phone Number & Enable Enterprise Voice
        # -----------------------------------------------------------------
        if (-not [string]::IsNullOrWhiteSpace($phone)) {
            # Normalize E.164 (strip whitespace/hyphens, preserve leading '+')
            $cleanPhone =$phone -replace '[\s\-\(\)]', ''
            if ($cleanPhone -notmatch '^\+') {
                $cleanPhone = "+$cleanPhone"
            }

            Write-Host "  [-] Assigning PSTN Phone Number: $cleanPhone ($PhoneNumberType)..." -ForegroundColor Gray
            if ($PSCmdlet.ShouldProcess($upn, "Assign Phone Number $cleanPhone")) {
                Set-CsPhoneNumberAssignment -Identity $upn `
                                            -PhoneNumber $cleanPhone `
                                            -PhoneNumberType $PhoneNumberType `
                                            -EnterpriseVoiceEnabled $true -ErrorAction Stop
                Write-Host "  [+] Phone number assigned and Enterprise Voice enabled." -ForegroundColor Green
            }
        } else {
            Write-Host "  [!] No telephone number found in CSV for this account. Skipping number assignment." -ForegroundColor DarkYellow
        }

        # -----------------------------------------------------------------
        # Step D: Apply Voice & International Calling Policies
        # -----------------------------------------------------------------
        Write-Host "  [-] Granting international dialing and device management policies..." -ForegroundColor Gray

        # 1. Teams Calling Policy (PSTN / International permission)
        if ($CallingPolicy) {
            Grant-CsTeamsCallingPolicy -Identity $upn -PolicyName $CallingPolicy -ErrorAction Stop
            Write-Host "      Granted Calling Policy: '$CallingPolicy'" -ForegroundColor Gray
        }

        # 2. Voice Routing Policy (Direct Routing / Operator Connect routing to international SBCs)
        if ($VoiceRoutingPolicy -and $PhoneNumberType -eq "DirectRouting") {
            Grant-CsOnlineVoiceRoutingPolicy -Identity $upn -PolicyName $VoiceRoutingPolicy -ErrorAction Stop
            Write-Host "      Granted Voice Routing Policy: '$VoiceRoutingPolicy'" -ForegroundColor Gray
        }

        # 3. Dial Plan (International normalization rules)
        if ($DialPlan -and $DialPlan -ne "Global") {
            Grant-CsTenantDialPlan -Identity $upn -PolicyName $DialPlan -ErrorAction Stop
            Write-Host "      Granted Dial Plan: '$DialPlan'" -ForegroundColor Gray
        }

        # 4. IP Phone Policy (Device lock timeout, shared console behavior)
        if ($IPPhonePolicy) {
            Grant-CsTeamsIPPhonePolicy -Identity $upn -PolicyName $IPPhonePolicy -ErrorAction Stop
            Write-Host "      Granted IP Phone Policy: '$IPPhonePolicy'" -ForegroundColor Gray
        }

        Write-Host "  [+] Account provisioning and policy alignment completed for $upn." -ForegroundColor Green
        $successCount++

    } catch {
        Write-Error "Failed to process user '$upn': $($_.Exception.Message)"
        $failCount++
    }
}

# -------------------------------------------------------------------------
# Execution Summary
# -------------------------------------------------------------------------
Write-Host "`n=======================================================" -ForegroundColor Cyan
Write-Host "           PROVISIONING SUMMARY" -ForegroundColor Cyan
Write-Host "=======================================================" -ForegroundColor Cyan
Write-Host " Successfully Processed : $successCount" -ForegroundColor Green
Write-Host " Failed / Skipped       : $failCount" -ForegroundColor $(if ($failCount -gt 0) { "Red" } else { "Gray" })
Write-Host "=======================================================" -ForegroundColor Cyan