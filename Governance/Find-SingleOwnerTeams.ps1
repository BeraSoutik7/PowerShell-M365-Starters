<#
================================================================================
SCRIPT NAME : Find-SingleOwnerTeams.ps1
DESCRIPTION : Scans tenant Microsoft Teams to identify governance risks:
              teams with only a single owner or zero owners (orphans).
AUTHOR      : Soutik Bera
REPOSITORY  : POWERSHELL-M365-STARTERS
MODULES REQ : MicrosoftTeams (v4.0.0+)
================================================================================
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false, HelpMessage = "Export report destination path")]
    [string]$OutputPath = ".\Governance\Teams_SingleOwner_Report_$(Get-Date -Format 'yyyyMMdd-HHmmss').csv",

    [Parameter(Mandatory = $false, HelpMessage = "Include orphaned teams with zero owners")]
    [switch]$IncludeZeroOwnerOrphans
)

$ErrorActionPreference = 'Stop'

# ==============================================================================
# SECTION 0: SESSION CONNECTION & PREREQUISITES
# ==============================================================================
Write-Host "[*] Checking Microsoft Teams session connection..." -ForegroundColor Cyan

try {
    $null = Get-CsTenant -ErrorAction Stop
}
catch {
    Write-Host "[*] Initiating connection to Microsoft Teams..." -ForegroundColor Yellow
    Connect-MicrosoftTeams
}

# ==============================================================================
# SECTION 1: DISCOVER AND AUDIT TEAMS
# ==============================================================================
Write-Host "`n[*] Discovering all Microsoft Teams in tenant..." -ForegroundColor Cyan
$AllTeams = Get-Team
Write-Host "[+] Found $($AllTeams.Count) total Teams. Starting owner governance audit..." -ForegroundColor Green

$AuditReport = [System.Collections.Generic.List[PSCustomObject]]::new()
$counter = 0

foreach ($team in $AllTeams) {$counter++
    Write-Progress -Activity "Auditing Team Ownership" `
                   -Status "Processing: $($team.DisplayName) ($counter of $($AllTeams.Count))" `
                   -PercentComplete (($counter / $AllTeams.Count) * 100)

    try {
        # Retrieve all owners assigned to the current team
        $owners = Get-TeamUser -GroupId $team.GroupId -Role Owner -ErrorAction Stop

        $ownerCount = if ($owners) {$owners.Count } else { 0 }

        # Check for Single Owner risk ($ownerCount -eq 1) or Orphaned risk ($ownerCount -eq 0)
        $isSingleOwner = ($ownerCount -eq 1)
        $isOrphan      = ($ownerCount -eq 0)

        if ($isSingleOwner -or ($IncludeZeroOwnerOrphans -and $isOrphan)) {$ownerUPN  = if ($ownerCount -eq 1) {$owners.User } elseif ($isOrphan) { "<NO_ACTIVE_OWNER>" } else { ($owners.User -join " ; ") }
            $ownerName = if ($ownerCount -eq 1) {$owners.Name } elseif ($isOrphan) { "<NO_ACTIVE_OWNER>" } else { ($owners.Name -join " ; ") }
            $riskState = if ($isOrphan) { "Orphaned (0 Owners)" } else { "Single Owner (1 Owner)" }

            $AuditReport.Add([PSCustomObject]@{
                TeamName       = $team.DisplayName
                GroupId        = $team.GroupId
                RiskLevel      = $riskState
                OwnerCount     = $ownerCount
                OwnerUPN       = $ownerUPN
                OwnerName      = $ownerName
                Archived       = $team.Archived
                ScanTimestamp  = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            })
        }
    }
    catch {
        Write-Warning "[-] Error auditing team '$($team.DisplayName)' ($($team.GroupId)):$_"
    }
}

Write-Progress -Activity "Auditing Team Ownership" -Completed

# ==============================================================================
# SECTION 2: OUTPUT & REPORTING
# ==============================================================================
Write-Host "`n[*] Audit Complete!" -ForegroundColor Cyan
Write-Host "--------------------------------------------------------"
Write-Host "Total Teams Scanned       : $($AllTeams.Count)"
Write-Host "At-Risk Teams Identified  : $($AuditReport.Count)" -ForegroundColor $(if ($AuditReport.Count -gt 0) { "Yellow" } else { "Green" })
Write-Host "--------------------------------------------------------`n"

if ($AuditReport.Count -gt 0) {
    # 2.1 Display summary table in terminal
    $AuditReport | Format-Table TeamName, RiskLevel, OwnerName, OwnerUPN -AutoSize

    # 2.2 Export persistent audit CSV
    $parentDir = Split-Path -Path $OutputPath -Parent
    if ($parentDir -and (-not (Test-Path -Path $parentDir))) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }

    $AuditReport | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8
    Write-Host "[+] Audit report exported successfully to: $OutputPath" -ForegroundColor Green
}
else {
    Write-Host "[+] All scanned teams comply with multi-ownership governance baseline." -ForegroundColor Green
}