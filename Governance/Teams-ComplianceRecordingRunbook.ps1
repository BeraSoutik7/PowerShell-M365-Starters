<#
================================================================================
SCRIPT NAME : Teams-ComplianceRecordingRunbook.ps1
DESCRIPTION : Modular command runbook for Microsoft Teams Compliance Recording 
              Policies (Create, Verify, Assign, Remove, and Delete).
AUTHOR      : Shautik Bera
REPOSITORY  : POWERSHELL-M365-STARTERS
MODULES REQ : MicrosoftTeams (v4.0.0+)
================================================================================
#>

# ==============================================================================
# SECTION 0: SESSION CONNECTION & PREREQUISITES
# ==============================================================================

# Install the Microsoft Teams module if not present
# Install-Module -Name MicrosoftTeams -Scope CurrentUser -Repository PSGallery -Force

# Authenticate to Microsoft Teams
Connect-MicrosoftTeams


# ==============================================================================
# SECTION 1: INSPECTION & VERIFICATION COMMANDS
# ==============================================================================

# 1.1 List all compliance recording policies defined in the tenant
Get-CsTeamsComplianceRecordingPolicy | Select-Object Identity, Description, Enabled | Format-Table -AutoSize

# 1.2 Inspect details of a specific policy
Get-CsTeamsComplianceRecordingPolicy -Identity "FinanceComplianceRecordingPolicy" | Format-List

# 1.3 Check assigned recording policy for a single user
$UserToCheck = "alex.wilber@contoso.com"
(Get-CsOnlineUser -Identity $UserToCheck) | Select-Object UserPrincipalName, DisplayName, TeamsComplianceRecordingPolicy | Format-Table -AutoSize

# 1.4 Check assigned recording policy for a bulk list of users via CSV
$CsvVerificationPath = ".\Governance\ComplianceRecording.sample.csv"
if (Test-Path $CsvVerificationPath) {
    Import-Csv -Path $CsvVerificationPath | ForEach-Object {
        $upn = if ($_.UserPrincipalName) { $_.UserPrincipalName.Trim() } else {$_.UserId.Trim() }
        try {
            $userObj = Get-CsOnlineUser -Identity$upn -ErrorAction Stop
            [PSCustomObject]@{
                UserPrincipalName          = $userObj.UserPrincipalName
                TeamsComplianceRecordingPolicy = if ($userObj.TeamsComplianceRecordingPolicy) {$userObj.TeamsComplianceRecordingPolicy } else { "<None / Global Default>" }
                Status                     = "Success"
            }
        }
        catch {
            [PSCustomObject]@{
                UserPrincipalName          = $upn
                TeamsComplianceRecordingPolicy = $null
                Status                     = "User Not Found / Error"
            }
        }
    } | Format-Table -AutoSize
} else {
    Write-Warning "File not found: $CsvVerificationPath"
}


# ==============================================================================
# SECTION 2: CREATE RECORDING POLICY
# ==============================================================================

$NewPolicyName = "FinanceComplianceRecordingPolicy"
$PolicyDesc    = "Mandatory compliance call recording policy for financial trading desks."

try {
    $existingPolicy = Get-CsTeamsComplianceRecordingPolicy -Identity$NewPolicyName -ErrorAction SilentlyContinue
    if (-not $existingPolicy) {
        New-CsTeamsComplianceRecordingPolicy `
            -Identity $NewPolicyName `
            -Description $PolicyDesc
        Write-Host "[+] Policy '$NewPolicyName' created successfully." -ForegroundColor Green
    } else {
        Write-Warning "[-] Policy '$NewPolicyName' already exists in this tenant."
    }
}
catch {
    Write-Error "Failed to create policy '$NewPolicyName':$_"
}


# ==============================================================================
# SECTION 3: ASSIGN POLICY (SINGLE USER & BULK CSV)
# ==============================================================================

# ----------------- 3.1 Single User Assignment -----------------
$TargetUser = "alex.wilber@contoso.com"
$TargetPolicy = "FinanceComplianceRecordingPolicy"

try {
    Grant-CsTeamsComplianceRecordingPolicy -Identity $TargetUser -PolicyName$TargetPolicy -ErrorAction Stop
    Write-Host "[+] Assigned policy '$TargetPolicy' to$TargetUser" -ForegroundColor Green
}
catch {
    Write-Error "[-] Failed to assign policy to $TargetUser:$_"
}

# ----------------- 3.2 Bulk User Assignment via CSV -----------
$CsvAssignPath = ".\Governance\ComplianceRecording.sample.csv"
$BulkPolicyName = "FinanceComplianceRecordingPolicy"

if (Test-Path $CsvAssignPath) {
    $UserBatch = Import-Csv -Path$CsvAssignPath
    Write-Host "`n[*] Starting bulk assignment for $($UserBatch.Count) users..." -ForegroundColor Cyan

    foreach ($entry in $UserBatch) {
        $upn = if ($entry.UserPrincipalName) { $entry.UserPrincipalName.Trim() } else { $entry.UserId.Trim() }
        
        if (-not [string]::IsNullOrWhiteSpace($upn)) {
            try {
                Grant-CsTeamsComplianceRecordingPolicy -Identity $upn -PolicyName $BulkPolicyName -ErrorAction Stop
                Write-Host "[+] Assigned: $upn" -ForegroundColor Green
            }
            catch {
                Write-Warning "[-] Failed for ${upn}: $_"
            }
        }
    }
} else {
    Write-Warning "CSV path '$CsvAssignPath' does not exist."
}


# ==============================================================================
# SECTION 4: UNASSIGN / REMOVE POLICY (REVERT TO GLOBAL DEFAULT)
# ==============================================================================

# ----------------- 4.1 Single User Removal --------------------
$UserToRevert = "alex.wilber@contoso.com"

try {
    # Passing $null unassigns explicit policy and reverts user to default
    Grant-CsTeamsComplianceRecordingPolicy -Identity $UserToRevert -PolicyName $null -ErrorAction Stop
    Write-Host "[+] Removed explicit policy from $UserToRevert (reverted to default)." -ForegroundColor Green
}
catch {
    Write-Error "[-] Failed to unassign policy for $UserToRevert: $_"
}

# ----------------- 4.2 Bulk Removal via CSV -------------------
$CsvRemovePath = ".\Governance\ComplianceRecording.sample.csv"

if (Test-Path $CsvRemovePath) {
    $RemoveBatch = Import-Csv -Path $CsvRemovePath
    Write-Host "`n[*] Starting bulk policy unassignment..." -ForegroundColor Cyan

    foreach ($entry in$RemoveBatch) {
        $upn = if ($entry.UserPrincipalName) { $entry.UserPrincipalName.Trim() } else {$entry.UserId.Trim() }

        if (-not [string]::IsNullOrWhiteSpace($upn)) {
            try {
                Grant-CsTeamsComplianceRecordingPolicy -Identity $upn -PolicyName$null -ErrorAction Stop
                Write-Host "[+] Removed policy from: $upn" -ForegroundColor Yellow
            }
            catch {
                Write-Warning "[-] Failed to revert ${upn}:$_"
            }
        }
    }
} else {
    Write-Warning "CSV path '$CsvRemovePath' does not exist."
}


# ==============================================================================
# SECTION 5: DECOMMISSION / DELETE POLICY ENTIRELY
# ==============================================================================

$PolicyToDelete = "FinanceComplianceRecordingPolicy"

try {
    # Ensure no users remain assigned before executing this command
    Remove-CsTeamsComplianceRecordingPolicy -Identity $PolicyToDelete -Confirm:$false -ErrorAction Stop
    Write-Host "[+] Policy '$PolicyToDelete' has been permanently deleted from tenant." -ForegroundColor Green
}
catch {
    Write-Error "[-] Failed to delete policy '$PolicyToDelete':$_"
}