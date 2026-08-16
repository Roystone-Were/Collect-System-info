<#
.SYNOPSIS
    Pushes all collected *_sysinfo.csv files into a SharePoint list.
    Run this on the IT lead's machine - NOT on each laptop.

.DESCRIPTION
    - Reads every *_sysinfo.csv in -Folder (default: .\Results)
    - Keeps the latest row per computer (dedupe by ComputerName)
    - Creates the SharePoint list + columns automatically if missing
    - The list's Title column shows the OWNER (typed at collection time),
      because computer names are not human-friendly
    - Upserts by ComputerName: re-running updates existing rows,
      it never creates duplicates
    - Long fields (disks, monitors, keyboard, mouse, scanner) are stored
      as multi-line text (single-line fields cap at 255 chars)

    ONE-TIME AUTH SETUP (PnP PowerShell needs an Entra app registration):
    If you have Global Admin rights, run once on the IT machine:
        Register-PnPEntraIDApp -DisplayName "IT Inventory Sync" -Interactive
    ...consent when asked, note the ClientId it prints, then always call
    this script with -ClientId <that id>.
    Without admin rights, try first without -ClientId; if login is refused,
    ask your M365 admin to run the command above once for the tenant.

.PARAMETER SiteUrl
    Your SharePoint site, e.g. "https://refrontier.sharepoint.com/sites/IT"

.PARAMETER ListName
    The inventory list name. Default: "IT Inventory". Created if missing.

.PARAMETER Folder
    Folder holding the *_sysinfo.csv files. Default: .\Results

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File .\Sync-ToSharePoint.ps1 `
        -SiteUrl "https://refrontier.sharepoint.com/sites/IT" -ClientId "guid"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,
    [string]$ListName = 'IT Inventory',
    [string]$Folder = '',
    [string]$ClientId = ''
)

$ErrorActionPreference = 'Stop'

# Resolve defaults in the body: $PSScriptRoot can be empty in param defaults on Windows PowerShell 5.1
if (-not $Folder) { $Folder = Join-Path $PSScriptRoot 'Results' }

#--------------------------------------------------------------
# 1) Ensure PnP PowerShell (per-user install, no admin needed)
#--------------------------------------------------------------
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host '[i] Installing PnP.PowerShell for the current user (one-time).' -ForegroundColor Cyan
    Write-Host '    Downloads ~90 MB. The first ~30 seconds can look quiet -' -ForegroundColor DarkGray
    Write-Host '    then a progress bar appears. Do not close this window.' -ForegroundColor DarkGray
    $ProgressPreference = 'Continue'   # ensure the download progress bar is displayed
    try {
        Install-Module PnP.PowerShell -Scope CurrentUser -Force -SkipPublisherCheck -ErrorAction Stop
    } catch {
        throw "Could not install PnP.PowerShell automatically. Run this once in a normal console and answer Yes:`n    Install-Module PnP.PowerShell -Scope CurrentUser -Force"
    }
}
Import-Module PnP.PowerShell

#--------------------------------------------------------------
# 2) Connect (interactive browser login)
#--------------------------------------------------------------
$connect = @{ Url = $SiteUrl; Interactive = $true }
if ($ClientId) { $connect['ClientId'] = $ClientId }
Connect-PnPOnline @connect
Write-Host "[OK] Connected to $SiteUrl" -ForegroundColor Green

#--------------------------------------------------------------
# 3) Load rows: latest per computer
#--------------------------------------------------------------
$files = @(Get-ChildItem -LiteralPath $Folder -Filter '*_sysinfo.csv' -File -ErrorAction SilentlyContinue)
if ($files.Count -eq 0) { Write-Warning "No *_sysinfo.csv files found in '$Folder'."; return }
$rows = @()
foreach ($f in $files) {
    try { $rows += Import-Csv -LiteralPath $f.FullName }
    catch { Write-Warning "Skipping unreadable file '$($f.Name)': $($_.Exception.Message)" }
}
$machines = @($rows |
    Sort-Object ComputerName, { [datetime]$_.Collected } |
    Group-Object ComputerName |
    ForEach-Object { $_.Group[-1] })
if ($machines.Count -eq 0) { Write-Warning 'No readable inventory rows found.'; return }
Write-Host "[i] $($machines.Count) machine(s) to sync (from $($files.Count) file(s))." -ForegroundColor Cyan

#--------------------------------------------------------------
# 4) Ensure the list + columns exist
#--------------------------------------------------------------
if (-not (Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue)) {
    New-PnPList -Title $ListName -Template GenericList | Out-Null
    Write-Host "[OK] Created list '$ListName'." -ForegroundColor Green
}
# Make the Title column display as "Owner" (Title = owner name, not computer name)
try { Set-PnPField -List $ListName -Identity 'Title' -Values @{ Title = 'Owner' } | Out-Null } catch {}

$longText   = @('Disks', 'Monitor(s)', 'Keyboard', 'Mouse', 'Scanner')   # >255 chars possible
$numberCols = @('Cores', 'Threads', 'TotalRAM_GB', 'RAM_Modules', 'RAM_Max_GB', 'RAM_FreeSlots', 'DriveC_Total_GB', 'DriveC_Free_GB', 'Uptime_Days', 'Defender_SigAge_Days', 'Defender_LastScan_Days')
$dateCols   = @('Collected')

$fieldMap = @{}   # CSV column name -> SharePoint field internal name
foreach ($col in ($machines[0].PSObject.Properties.Name)) {
    $internal = ($col -replace '[^A-Za-z0-9_]', '')
    if ($internal -eq 'Owner') { continue }          # Owner goes into the Title column
    $fieldMap[$col] = $internal
    if (-not (Get-PnPField -List $ListName -Identity $internal -ErrorAction SilentlyContinue)) {
        $type = if ($dateCols -contains $col) { 'DateTime' }
                elseif ($numberCols -contains $col) { 'Number' }
                elseif ($longText -contains $col) { 'MultiLineText' }
                else { 'Text' }
        Add-PnPField -List $ListName -InternalName $internal -DisplayName $col -Type $type -AddToDefaultView | Out-Null
        Write-Host "    + column '$col' ($type)"
    }
}
# ComputerName is the upsert key - index it so lookups stay fast
Set-PnPField -List $ListName -Identity 'ComputerName' -Values @{ Indexed = $true } | Out-Null

#--------------------------------------------------------------
# 5) Upsert: update existing machine rows, add new ones
#--------------------------------------------------------------
$added = 0; $updated = 0; $failed = 0
foreach ($m in $machines) {
    try {
        $cnFilter = $m.ComputerName -replace "'", "''"
        $existing = @(Get-PnPListItem -List $ListName -Filter "ComputerName eq '$cnFilter'")

        $values = @{ 'Title' = [string]$m.Owner }
        foreach ($col in $fieldMap.Keys) {
            $v = $m.$col
            if ($dateCols -contains $col -and $v) {
                $values[$fieldMap[$col]] = [datetime]$v
            } elseif ($numberCols -contains $col) {
                $n = 0.0
                if ([double]::TryParse([string]$v, [Globalization.CultureInfo]::InvariantCulture, [ref]$n)) {
                    $values[$fieldMap[$col]] = $n
                }
            } else {
                $values[$fieldMap[$col]] = [string]$v
            }
        }

        if ($existing.Count -gt 0) {
            Set-PnPListItem -List $ListName -Identity $existing[0].Id -Values $values | Out-Null
            $updated++
        } else {
            Add-PnPListItem -List $ListName -Values $values | Out-Null
            $added++
        }
    } catch {
        Write-Warning "Failed: $($m.ComputerName) - $($_.Exception.Message)"
        $failed++
    }
}
Write-Host "[OK] Sync complete: $added added, $updated updated, $failed failed." -ForegroundColor Green
