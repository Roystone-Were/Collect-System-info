<#
.SYNOPSIS
    Merge all *_sysinfo.csv files collected from laptops into ONE Excel workbook
    and a master CSV. Run on the IT lead's machine (NOT on each laptop).

.DESCRIPTION
    - Imports every *_sysinfo.csv in the folder
    - Skips files that fail to import (warns, does not stop)
    - Keeps the latest row per computer (dedupe by ComputerName)
    - Writes Master_Inventory.csv
    - Writes Master_Inventory.xlsx (via the free ImportExcel module).
      If ImportExcel is missing, installs it (one-time) or falls back to CSV.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File Merge-ITInventory.ps1
.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File Merge-ITInventory.ps1 -Folder "C:\Inventory"
#>

[CmdletBinding()]
param(
    [string]$Folder = '',
    [string]$OutCsv   = 'Master_Inventory.csv',
    [string]$OutExcel = 'Master_Inventory.xlsx'
)

$ErrorActionPreference = 'Stop'

# Resolve default in the body: $PSScriptRoot can be empty in param defaults on Windows PowerShell 5.1
if (-not $Folder) { $Folder = $PSScriptRoot }

if (-not (Test-Path -LiteralPath $Folder)) {
    Write-Error "Folder not found: $Folder"
    return
}

$files = @(Get-ChildItem -LiteralPath $Folder -Filter '*_sysinfo.csv' -File -ErrorAction SilentlyContinue)
if ($files.Count -eq 0) {
    Write-Warning "No *_sysinfo.csv files found in '$Folder'."
    return
}

# Combine rows, skipping any file that cannot be imported
$rows = foreach ($f in $files) {
    try {
        Import-Csv -LiteralPath $f.FullName
    } catch {
        Write-Warning "Skipping unreadable file '$($f.Name)': $($_.Exception.Message)"
    }
}
$rows = @($rows)

# Keep the LATEST entry per computer
$deduped = $rows |
    Sort-Object ComputerName, Collected |
    Group-Object ComputerName |
    ForEach-Object { $_.Group[-1] }
$deduped = @($deduped)

$csvPath = Join-Path $Folder $OutCsv
$deduped | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "[OK] Master CSV:   $csvPath   ($($deduped.Count) machines from $($files.Count) files)" -ForegroundColor Green

# Excel workbook via ImportExcel
$excelDone = $false
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "[i] Installing ImportExcel module (one-time, from PSGallery)..." -ForegroundColor Cyan
    try { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue } catch {}
    try { Install-PackageProvider -Name NuGet -Force -Scope CurrentUser -ErrorAction Stop } catch {}
    try {
        Install-Module ImportExcel -Scope CurrentUser -Force -Confirm:$false -AllowClobber -SkipPublisherCheck -ErrorAction Stop
    } catch {
        Write-Warning "Could not install ImportExcel: $($_.Exception.Message)"
        Write-Host "    Run this once in a NORMAL console window (not automated) and answer Yes:" -ForegroundColor Yellow
        Write-Host "    Install-Module ImportExcel -Scope CurrentUser -Force" -ForegroundColor Cyan
    }
}
if (Get-Module -ListAvailable -Name ImportExcel) {
    try {
        Import-Module ImportExcel -ErrorAction Stop
        $xlsxPath = Join-Path $Folder $OutExcel
        $deduped | Export-Excel -Path $xlsxPath -AutoSize -TableName 'Inventory' -WorksheetName 'Inventory'
        Write-Host "[OK] Excel workbook: $xlsxPath" -ForegroundColor Green
        $excelDone = $true
    } catch {
        Write-Warning "Excel export failed: $($_.Exception.Message)"
    }
}

if (-not $excelDone) {
    Write-Host "[!] No .xlsx was created. The master CSV still opens in Excel." -ForegroundColor Yellow
    Write-Host "    To get a native .xlsx later, run on your machine:" -ForegroundColor Yellow
    Write-Host "    Install-Module ImportExcel -Scope CurrentUser -Force" -ForegroundColor Cyan
    Write-Host "    .\Merge-ITInventory.ps1" -ForegroundColor Cyan
}
