<#
.SYNOPSIS
    Updates and verifies Microsoft 365 UsageLocation for single or bulk users.

.DESCRIPTION
    This script connects to Microsoft Graph and Microsoft Teams PowerShell modules
    to update user usage locations individually or in bulk via a CSV file, then
    verifies the assigned UsageLocation.

.NOTES
    Author: Soutik Bera
    Requires: Microsoft.Graph, MicrosoftTeams PowerShell modules
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false, HelpMessage = "Single UserPrincipalName or Object ID")]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $false, HelpMessage = "Path to CSV file containing users")]
    [string]$CsvPath,

    [Parameter(Mandatory = $false, HelpMessage = "Two-letter ISO country code (e.g., US, GB, IN)")]
    [ValidateLength(2, 2)]
    [string]$UsageLocation = "US",

    [Parameter(Mandatory = $false, HelpMessage = "Switch to verify usage location using Teams module")]
    [switch]$VerifyTeams
)

# ---------------------------------------------------------
# Connect to Microsoft Graph
# ---------------------------------------------------------
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.ReadWrite.All" -NoWelcome

# ---------------------------------------------------------
# Scenario 1: Update Single User
# ---------------------------------------------------------
if ($UserPrincipalName) {
    try {
        Write-Host "Updating UsageLocation to '$UsageLocation' for: $UserPrincipalName" -ForegroundColor Yellow
        Update-MgUser -UserId $UserPrincipalName -UsageLocation $UsageLocation
        Write-Host "Successfully updated $UserPrincipalName" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to update $UserPrincipalName: $_"
    }
}

# ---------------------------------------------------------
# Scenario 2: Bulk Update via CSV
# ---------------------------------------------------------
if ($CsvPath) {
    if (Test-Path $CsvPath) {
        $users = Import-Csv -Path $CsvPath
        Write-Host "Imported $($users.Count) users from $CsvPath" -ForegroundColor Cyan

        foreach ($user in $users) {
            # Assumes CSV contains a header named 'UserPrincipalName' or 'UserId'
            $targetUser = if ($user.UserPrincipalName) { $user.UserPrincipalName } else { $user.UserId }

            if (-not [string]::IsNullOrWhiteSpace($targetUser)) {
                try {
                    Update-MgUser -UserId $targetUser.Trim() -UsageLocation $UsageLocation
                    Write-Host "Updated: $targetUser -> $UsageLocation" -ForegroundColor Green
                }
                catch {
                    Write-Warning "Error updating ${targetUser}: $_"
                }
            }
        }
    }
    else {
        Write-Error "CSV file not found at: $CsvPath"
    }
}

# ---------------------------------------------------------
# Scenario 3: Verification (Optional via Teams module)
# ---------------------------------------------------------
if ($VerifyTeams -and $CsvPath -and (Test-Path $CsvPath)) {
    Write-Host "`nVerifying Usage Locations via Teams/Skype module..." -ForegroundColor Cyan

    $userList = Import-Csv -Path $CsvPath
    foreach ($entry in $userList) {
        $id = if ($entry.UserPrincipalName) { $entry.UserPrincipalName.Trim() } else { $entry.UserId.Trim() }

        if (-not [string]::IsNullOrWhiteSpace($id)) {
            try {
                $details = Get-CsOnlineUser -Identity $id
                Write-Host "$($details.UserPrincipalName): $($details.UsageLocation)" -ForegroundColor Green
            }
            catch {
                Write-Host "$id: Not found or error retrieving details" -ForegroundColor Yellow
            }
        }
    }
}