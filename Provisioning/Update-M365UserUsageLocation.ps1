<#
================================================================================
SCRIPT NAME : M365-UsageLocationRunbook.ps1
DESCRIPTION : Modular command runbook for Microsoft 365 User UsageLocation
              (Inspect, Update Single, Update Bulk, Audit, and Clear).
AUTHOR      : Soutik Bera
REPOSITORY  : POWERSHELL-M365-STARTERS
MODULES REQ : Microsoft.Graph.Users, MicrosoftTeams (Optional for Voice Audit)
================================================================================
#>

# ==============================================================================
# SECTION 0: SESSION CONNECTION & PREREQUISITES
# ==============================================================================

# Install the required modules if not present
# Install-Module -Name Microsoft.Graph.Users -Scope CurrentUser -Repository PSGallery -Force
# Install-Module -Name MicrosoftTeams -Scope CurrentUser -Repository PSGallery -Force

# Authenticate to Microsoft Graph with Directory User management permissions
Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All" -NoWelcome

# (Optional) Authenticate to Microsoft Teams to verify directory replication
Connect-MicrosoftTeams


# ==============================================================================
# SECTION 1: INSPECTION & VERIFICATION COMMANDS
# ==============================================================================

# 1.1 Check UsageLocation for a single user in Microsoft Graph
$UserToCheck = "alex.wilber@contoso.com"
(Get-MgUser -UserId $UserToCheck -Property DisplayName, UserPrincipalName, UsageLocation) | 
    Select-Object DisplayName, UserPrincipalName, UsageLocation | Format-Table -AutoSize

# 1.2 Cross-verify user UsageLocation downstream in Microsoft Teams Online
Get-CsOnlineUser -Identity $UserToCheck | 
    Select-Object DisplayName, UserPrincipalName, UsageLocation, EnterpriseVoiceEnabled | Format-Table -AutoSize

# 1.3 Check assigned UsageLocation for a bulk list of users via CSV
$CsvVerificationPath = ".\Governance\UsageLocation.sample.csv"
if (Test-Path $CsvVerificationPath) {
    Import-Csv -Path $CsvVerificationPath | ForEach-Object {
        $upn = if ($_.UserPrincipalName) { $_.UserPrincipalName.Trim() } else {$_.UserId.Trim() }
        try {
            $graphUser = Get-MgUser -UserId$upn -Property DisplayName, UserPrincipalName, UsageLocation -ErrorAction Stop
            [PSCustomObject]@{
                UserPrincipalName = $graphUser.UserPrincipalName
                DisplayName       = $graphUser.DisplayName
                UsageLocation     = if ($graphUser.UsageLocation) {$graphUser.UsageLocation } else { "<Not Set>" }
                Status            = "Success"
            }
        }
        catch {
            [PSCustomObject]@{
                UserPrincipalName = $upn
                DisplayName       = $null
                UsageLocation     = $null
                Status            = "User Not Found / Error"
            }
        }
    } | Format-Table -AutoSize
} else {
    Write-Warning "File not found: $CsvVerificationPath"
}


# ==============================================================================
# SECTION 2: UPDATE USAGELOCATION (SINGLE USER)
# ==============================================================================

$TargetUser     = "alex.wilber@contoso.com"
$TargetLocation = "US" # Two-letter ISO country code (e.g., US, GB, IN, CA)

try {
    Update-MgUser -UserId $TargetUser -UsageLocation$TargetLocation -ErrorAction Stop
    Write-Host "[+] Successfully updated UsageLocation to '$TargetLocation' for$TargetUser" -ForegroundColor Green
}
catch {
    Write-Error "[-] Failed to update UsageLocation for $TargetUser:$_"
}


# ==============================================================================
# SECTION 3: BULK UPDATE USAGELOCATION VIA CSV
# ==============================================================================

$CsvAssignPath   = ".\Governance\UsageLocation.sample.csv"
$BulkCountryCode = "US"

if (Test-Path $CsvAssignPath) {
    $UserBatch = Import-Csv -Path$CsvAssignPath
    Write-Host "`n[*] Starting bulk UsageLocation assignment for $($UserBatch.Count) users..." -ForegroundColor Cyan

    foreach ($entry in$UserBatch) {
        $upn = if ($entry.UserPrincipalName) { $entry.UserPrincipalName.Trim() } else {$entry.UserId.Trim() }
        
        # Optionally allow per-row country override if 'UsageLocation' column exists in CSV
        $country = if ($entry.UsageLocation) { $entry.UsageLocation.Trim().ToUpper() } else {$BulkCountryCode }

        if (-not [string]::IsNullOrWhiteSpace($upn)) {
            try {
                Update-MgUser -UserId $upn -UsageLocation$country -ErrorAction Stop
                Write-Host "[+] Updated: $upn ->$country" -ForegroundColor Green
            }
            catch {
                Write-Warning "[-] Failed for ${upn}:$_"
            }
        }
    }
} else {
    Write-Warning "CSV path '$CsvAssignPath' does not exist."
}


# ==============================================================================
# SECTION 4: BULK VERIFICATION & AUDIT REPORT EXPORT
# ==============================================================================

$CsvAuditPath  = ".\Governance\UsageLocation.sample.csv"
$ReportOutPath = ".\Governance\UsageLocation_AuditReport.csv"

if (Test-Path $CsvAuditPath) {
    Write-Host "`n[*] Generating cross-platform directory audit report..." -ForegroundColor Cyan
    
    $AuditData = Import-Csv -Path $CsvAuditPath | ForEach-Object {
        $upn = if ($_.UserPrincipalName) { $_.UserPrincipalName.Trim() } else { $_.UserId.Trim() }

        if (-not [string]::IsNullOrWhiteSpace($upn)) {
            $graphUser = Get-MgUser -UserId $upn -Property UsageLocation, AccountEnabled -ErrorAction SilentlyContinue
            $teamsUser = Get-CsOnlineUser -Identity $upn -ErrorAction SilentlyContinue

            [PSCustomObject]@{
                UserPrincipalName   = $upn
                AccountEnabled      = if ($graphUser) { $graphUser.AccountEnabled } else { $false }
                GraphUsageLocation  = if ($graphUser.UsageLocation) { $graphUser.UsageLocation } else { "<Not Set>" }
                TeamsUsageLocation  = if ($teamsUser.UsageLocation) { $teamsUser.UsageLocation } else { "<Pending Sync>" }
                TeamsProvisioned    = [bool]$teamsUser
                AuditTimestamp      = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            }
        }
    }

    # Output to screen
    $AuditData | Format-Table -AutoSize

    # Export report to CSV
    $AuditData | Export-Csv -Path $ReportOutPath -NoTypeInformation
    Write-Host "[+] Audit report exported to: $ReportOutPath" -ForegroundColor Green
} else {
    Write-Warning "CSV path '$CsvAuditPath' does not exist."
}


# ==============================================================================
# SECTION 5: CLEAR / RESET USAGELOCATION
# ==============================================================================

# NOTE: Clearing UsageLocation in Microsoft Entra ID is only permitted 
# if the account currently has NO active licenses assigned.

# ----------------- 5.1 Single User Reset ----------------------
$UserToReset = "alex.wilber@contoso.com"

try {
    # Passing empty string clears the attribute if no license dependencies exist
    Update-MgUser -UserId $UserToReset -UsageLocation "" -ErrorAction Stop
    Write-Host "[+] Cleared UsageLocation for $UserToReset" -ForegroundColor Yellow
}
catch {
    Write-Error "[-] Failed to clear UsageLocation for $UserToReset (Check for active licenses): $_"
}

# ----------------- 5.2 Bulk User Reset via CSV ----------------
$CsvResetPath = ".\Governance\UsageLocation.sample.csv"

if (Test-Path $CsvResetPath) {
    $ResetBatch = Import-Csv -Path $CsvResetPath
    Write-Host "`n[*] Starting bulk UsageLocation clear operation..." -ForegroundColor Cyan

    foreach ($entry in$ResetBatch) {
        $upn = if ($entry.UserPrincipalName) { $entry.UserPrincipalName.Trim() } else {$entry.UserId.Trim() }

        if (-not [string]::IsNullOrWhiteSpace($upn)) {
            try {
                Update-MgUser -UserId $upn -UsageLocation "" -ErrorAction Stop
                Write-Host "[+] Cleared UsageLocation: $upn" -ForegroundColor Yellow
            }
            catch {
                Write-Warning "[-] Failed to clear ${upn}:$_"
            }
        }
    }
} else {
    Write-Warning "CSV path '$CsvResetPath' does not exist."
}