<#
.SYNOPSIS
    Fleet IT inventory collector for Windows 10 / Windows 11 laptops.
    Run once on EACH laptop. No admin rights required. Logs the row to a CSV.

.DESCRIPTION
    Collects:
      - PC name, logged-in username, detected WORK EMAIL
      - Windows edition/build, OS architecture
      - Manufacturer, model, serial number
      - CPU (name, cores, threads)
      - RAM (total GB, module count, speed)
      - Disks (SSD/HDD models + sizes), C: drive total/free space
      - IPv4 address, MAC address, GPU, battery presence, TPM status
      - Last boot time and uptime in days

    Output: <ComputerName>_sysinfo.csv  in the OutputFolder.
    Merge all CSV files into one Excel workbook using Merge-ITInventory.ps1.

.PARAMETER OutputFolder
    Where the CSV is saved. Default: Desktop\ITInventory
    For 50+ laptops, point this at a shared/network folder, e.g.:
    -OutputFolder "\\FILESERVER\Share\ITInventory"

.PARAMETER Master
    Also append this machine to a Master_Inventory.csv (deduped by computer name).

.PARAMETER IncludeOutlookCheck
    Optionally ask Outlook for the primary SMTP address (can be slow / may
    launch Outlook). Not required for most machines.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File Get-ITInventory.ps1
.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File Get-ITInventory.ps1 -OutputFolder "\\server\share\ITInventory" -Master
#>

[CmdletBinding()]
param(
    [string]$OutputFolder = (Join-Path ([Environment]::GetFolderPath('Desktop')) 'ITInventory'),
    [switch]$Master,
    [switch]$IncludeOutlookCheck,
    [string]$Owner = '',
    [switch]$AskOwner
)

$ErrorActionPreference = 'Continue'

#--------------------------------------------------------------
# Work email detection (tries several sources, returns first hit)
#--------------------------------------------------------------
function Get-WorkEmail {
    param([switch]$IncludeOutlook)

    # 1) Domain UPN (domain-joined machines)
    try {
        $upn = ((& whoami /upn 2>$null) -join '').Trim()
        if ($upn -match '\S+@\S+') { return $upn }
    } catch {}

    # 2) Microsoft 365 / Office identity stored in registry
    try {
        $base = 'HKCU:\Software\Microsoft\Office\16.0\Common\Identity\Identities'
        foreach ($key in (Get-ChildItem $base -ErrorAction SilentlyContinue)) {
            $ip = Get-ItemProperty -Path $key.PSPath -ErrorAction SilentlyContinue
            foreach ($field in 'UserPrincipalName','CommonName','Email') {
                if ($ip.$field -and $ip.$field -match '\S+@\S+') { return [string]$ip.$field }
            }
        }
    } catch {}

    # 2b) OneDrive work/business account (very reliable source of the work email)
    try {
        $odBase = 'HKCU:\Software\Microsoft\OneDrive\Accounts'
        foreach ($acct in (Get-ChildItem $odBase -ErrorAction SilentlyContinue)) {
            if ($acct.PSChildName -match '^Business') {
                $ip = Get-ItemProperty -Path $acct.PSPath -ErrorAction SilentlyContinue
                if ($ip.UserEmail -and $ip.UserEmail -match '\S+@\S+') { return [string]$ip.UserEmail }
            }
        }
    } catch {}

    # 2c) OneDrive personal / any account (fallback)
    try {
        $odBase = 'HKCU:\Software\Microsoft\OneDrive\Accounts'
        foreach ($acct in (Get-ChildItem $odBase -ErrorAction SilentlyContinue)) {
            $ip = Get-ItemProperty -Path $acct.PSPath -ErrorAction SilentlyContinue
            if ($ip.UserEmail -and $ip.UserEmail -match '\S+@\S+') { return [string]$ip.UserEmail }
        }
    } catch {}

    # 3) Microsoft account (IdentityCRL)
    try {
        foreach ($key in (Get-ChildItem 'HKCU:\Software\Microsoft\IdentityCRL\UserExtendedProperties' -ErrorAction SilentlyContinue)) {
            $ip = Get-ItemProperty -Path $key.PSPath -ErrorAction SilentlyContinue
            $hit = $ip.PSObject.Properties | Where-Object {
                $_.Name -notmatch '^PS' -and $_.Value -is [string] -and $_.Value -match '\S+@\S+'
            } | Select-Object -First 1
            if ($hit) { return [string]$hit.Value }
        }
    } catch {}

    # 4) Outlook primary SMTP address (only if requested)
    if ($IncludeOutlook) {
        try {
            $outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop
            $session  = $outlook.Session
            $address  = $session.CurrentUser.AddressEntry
            $smtp     = $null
            if ($address) {
                $exUser = $address.GetExchangeUser()
                if ($exUser) { $smtp = $exUser.PrimarySmtpAddress }
                elseif ($address.SMTPAddress) { $smtp = $address.SMTPAddress }
            }
            $outlook.Quit()
            if ($smtp -and $smtp -match '\S+@\S+') { return [string]$smtp }
        } catch {}
    }

    # 5) Fallback: domain\user hint
    if ($env:USERDNSDOMAIN) { return "$env:USERNAME@$env:USERDNSDOMAIN" }
    return 'Not found'
}

#--------------------------------------------------------------
# System info collection
#--------------------------------------------------------------
$computerName = $env:COMPUTERNAME
$userName     = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$workEmail    = Get-WorkEmail -IncludeOutlook:$IncludeOutlookCheck
# Owner: manually entered (when -AskOwner / -Owner given), else auto from the account
$ownerName = $userName
if ($Owner) { $ownerName = $Owner }
elseif ($AskOwner) {
    try {
        $reply = Read-Host 'Enter owner / name for THIS PC (ENTER = auto-detect)'
        if ($reply -and $reply.Trim()) { $ownerName = $reply.Trim() }
    } catch {}
}

$domain       = $env:USERDNSDOMAIN
if (-not $domain) { $domain = '(workgroup)' }

$os   = Get-CimInstance Win32_OperatingSystem
$cs   = Get-CimInstance Win32_ComputerSystem
$bios = Get-CimInstance Win32_BIOS
$prod = Get-CimInstance Win32_ComputerSystemProduct
$cpu  = Get-CimInstance Win32_Processor | Select-Object -First 1

$osName      = ($os.Caption -replace 'Microsoft Windows ','').Trim()
$osArch      = $os.OSArchitecture
$osBuild     = '{0} (10.0.{1})' -f $os.Version, $os.BuildNumber
$manufacturer= $cs.Manufacturer
$model       = $cs.Model
$serial      = $bios.SerialNumber
$cpuName     = if ($cpu.Name) { $cpu.Name.Trim() } else { 'N/A' }
$cores       = if ($cpu.NumberOfCores) { $cpu.NumberOfCores } else { 'N/A' }
$threads     = if ($cpu.NumberOfLogicalProcessors) { $cpu.NumberOfLogicalProcessors } else { 'N/A' }

# RAM
$totalRAM_GB = [math]::Round($cs.TotalPhysicalMemory / 1GB, 1)
$modules     = @(Get-CimInstance Win32_PhysicalMemory)
$ramModules  = $modules.Count
$ramSpeed    = if ($modules.Count) { ($modules.Speed -join '/') } else { 'N/A' }
# RAM upgrade info: max supported + free slots (Win32_PhysicalMemoryArray)
$ramArray    = @(Get-CimInstance Win32_PhysicalMemoryArray -ErrorAction SilentlyContinue)
$ramMaxKB    = ($ramArray | Measure-Object MaxCapacityEx -Sum).Sum
$ramMaxGB    = if ($ramMaxKB) { [math]::Round($ramMaxKB / 1MB, 0) } else { 'N/A' }
$ramSlots    = ($ramArray | Measure-Object MemoryDevices -Sum).Sum
$ramFreeSlot = if ($ramSlots) { [math]::Max(0, $ramSlots - $ramModules) } else { 'N/A' }

# Disks (SSD/HDD)
$diskList = @()
try {
    foreach ($d in (Get-PhysicalDisk -ErrorAction Stop)) {
        $sizeGB = [math]::Round($d.Size / 1GB, 0)
        $modelD = if ($d.Model) { $d.Model.Trim() } else { 'Unknown' }
        $diskList += '{0}  {1}GB  [{2}]' -f $modelD, $sizeGB, $d.MediaType
    }
} catch {
    foreach ($d in (Get-CimInstance Win32_DiskDrive)) {
        $sizeGB = [math]::Round($d.Size / 1GB, 0)
        $modelD = if ($d.Model) { $d.Model.Trim() } else { 'Unknown' }
        $diskList += '{0}  {1}GB' -f $modelD, $sizeGB
    }
}
$diskSummary = if ($diskList.Count) { $diskList -join '  |  ' } else { 'N/A' }

# Disk health (early warning for failing SSDs/HDDs)
$diskHealth = 'N/A'
try {
    $diskHealth = (Get-PhysicalDisk -ErrorAction Stop |
        ForEach-Object { '{0}:{1}' -f $_.FriendlyName, $_.HealthStatus }) -join '  |  '
    if (-not $diskHealth) { $diskHealth = 'N/A' }
} catch {}

# C: drive
$cDrive  = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
$cTotal  = if ($cDrive.Size)       { [math]::Round($cDrive.Size / 1GB, 1) }       else { 'N/A' }
$cFree   = if ($cDrive.FreeSpace)  { [math]::Round($cDrive.FreeSpace / 1GB, 1) }  else { 'N/A' }

# Network
$netCfg  = Get-NetIPConfiguration -ErrorAction SilentlyContinue |
           Where-Object { $_.IPv4DefaultGateway -and $_.NetAdapter.Status -eq 'Up' } |
           Select-Object -First 1
$ipAddr  = if ($netCfg) { $netCfg.IPv4Address.IPAddress } else { 'N/A' }
$macAddr = if ($netCfg) { $netCfg.NetAdapter.MacAddress }  else { 'N/A' }

# GPU / battery / TPM / uptime
$gpuName  = (Get-CimInstance Win32_VideoController | Select-Object -First 1).Name
$battery  = Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue | Select-Object -First 1
$batteryP = if ($battery) { 'Yes' } else { 'No' }
$tpmOn    = 'N/A'
try {
    $tpm = Get-CimInstance -Namespace 'root\cimv2\Security\MicrosoftTpm' -ClassName Win32_Tpm -ErrorAction Stop
    if ($tpm) { $tpmOn = if ($tpm.IsEnabled_InitialValue) { 'Yes' } else { 'No' } }
} catch {}
$lastBoot = if ($os.LastBootUpTime) { $os.LastBootUpTime.ToString('yyyy-MM-dd HH:mm') } else { 'N/A' }
$uptime   = if ($os.LastBootUpTime) { [math]::Round(((Get-Date) - $os.LastBootUpTime).TotalDays, 1) } else { 'N/A' }

#--------------------------------------------------------------
# Health & security posture
#--------------------------------------------------------------
# Battery wear: current full-charge capacity vs design capacity
$battHealth = 'N/A'
try {
    $battDesign = (Get-CimInstance -Namespace root\WMI -ClassName BatteryStaticData -ErrorAction Stop | Select-Object -First 1).DesignedCapacity
    $battFull   = (Get-CimInstance -Namespace root\WMI -ClassName BatteryFullChargedCapacity -ErrorAction Stop | Select-Object -First 1).FullChargedCapacity
    if ($battDesign -and $battFull) { $battHealth = '{0}%' -f [math]::Round(100 * $battFull / $battDesign, 0) }
} catch {}

# Windows Defender (absence usually means a 3rd-party AV is installed)
$defender = 'Not found (3rd-party AV?)'
$defSigAge = 'N/A'; $defScanAge = 'N/A'
try {
    $mp = Get-MpComputerStatus -ErrorAction Stop
    $defender = 'AV:{0} RT:{1}' -f $(if ($mp.AntivirusEnabled) {'On'} else {'Off'}),
                              $(if ($mp.RealTimeProtectionEnabled) {'On'} else {'Off'})
    if ($mp.AntivirusSignatureLastUpdated -and $mp.AntivirusSignatureLastUpdated.Year -gt 2000) {
        $defSigAge = [math]::Round(((Get-Date) - $mp.AntivirusSignatureLastUpdated).TotalDays, 0)
    }
    if ($mp.QuickScanEndTime -and $mp.QuickScanEndTime.Year -gt 2000) {
        $defScanAge = [math]::Round(((Get-Date) - $mp.QuickScanEndTime).TotalDays, 0)
    }
} catch {}

# Firewall per profile
$firewall = 'N/A'
try {
    $fw = Get-NetFirewallProfile -ErrorAction Stop
    if ($fw) { $firewall = ($fw | ForEach-Object { '{0}:{1}' -f $_.Name, $(if ($_.Enabled) {'On'} else {'Off'}) }) -join ' ' }
} catch {}

# Pending reboot (updates waiting for a restart)
$pendingReboot = 'No'
if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $pendingReboot = 'Yes' }
elseif (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $pendingReboot = 'Yes' }
elseif (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue) { $pendingReboot = 'Yes' }

# Windows activation / license status
$actMap = @{ 0='Unlicensed'; 1='Licensed'; 2='Grace'; 3='Grace'; 4='Non-genuine'; 5='Notification'; 6='Grace' }
$licStatus = (Get-CimInstance -ClassName SoftwareLicensingProduct -Filter "PartialProductKey IS NOT NULL AND Name LIKE 'Windows%'" -ErrorAction SilentlyContinue |
    Select-Object -First 1).LicenseStatus
if ($null -eq $licStatus) {
    $licStatus = (Get-CimInstance -ClassName SoftwareLicensingService -ErrorAction SilentlyContinue | Select-Object -First 1).LicenseStatus
}
$activation = if ($null -ne $licStatus -and $actMap.ContainsKey([int]$licStatus)) { $actMap[[int]$licStatus] } else { 'N/A' }

# BitLocker (reading the encryption namespace requires an elevated session)
$bitlocker = 'Run as admin to check'
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if ($isAdmin) {
    try {
        $vol = Get-CimInstance -Namespace root\cimv2\Security\MicrosoftVolumeEncryption -ClassName Win32_EncryptableVolume -Filter "DriveLetter='C:'" -ErrorAction Stop
        if ($vol) {
            $bitlocker = switch ([int]$vol.ProtectionStatus) { 0 {'Not encrypted'} 1 {'Encrypted'} 2 {'Protection off'} default {'Unknown'} }
        }
    } catch { $bitlocker = 'N/A' }
}

# OS support status (Windows 10 passed end of support in Oct 2025; Win11 builds are 22000+)
$osSupport = if ([int]$os.BuildNumber -lt 22000) { 'EOL - unsupported' } else { 'Supported' }

#--------------------------------------------------------------
# Peripherals: monitors, keyboard, mouse, scanner, other USB
#--------------------------------------------------------------
function ConvertFrom-Bytes {
    param($bytes)
    if ($null -eq $bytes) { return '' }
    return ([System.Text.Encoding]::ASCII.GetString($bytes)).Trim([char]0, ' ')
}

# Monitors (EDID via WMI - gives serial + model where available)
$monitorList = @()
Get-CimInstance -Namespace root\WMI -ClassName WmiMonitorID -ErrorAction SilentlyContinue |
    ForEach-Object {
        $m  = (ConvertFrom-Bytes $_.ManufacturerName).Trim()
        $nm = (ConvertFrom-Bytes $_.UserFriendlyName).Trim()
        $sn = (ConvertFrom-Bytes $_.SerialNumberID).Trim()
        $monitorList += ('{0} {1}  SN:{2}' -f $m, $nm, $sn).Trim()
    }
if ($monitorList.Count -eq 0) {
    Get-CimInstance Win32_DesktopMonitor -ErrorAction SilentlyContinue |
        Where-Object { $_.Name } | ForEach-Object { $monitorList += $_.Name }
}
$monitors = if ($monitorList.Count) { $monitorList -join '  |  ' } else { 'None detected' }

# Keyboard & mouse (with PNP device ID = VID/PID)
$keyboard = @(Get-CimInstance Win32_PnPEntity -ErrorAction SilentlyContinue |
    Where-Object { $_.PNPClass -eq 'Keyboard' -and $_.Status -eq 'OK' } |
    Select-Object -First 2)
$keyboardStr = if ($keyboard.Count) { ($keyboard | ForEach-Object { '{0}  [{1}]' -f $_.Name, $_.PNPDeviceID }) -join '  |  ' } else { 'None' }

$mouse = @(Get-CimInstance Win32_PnPEntity -ErrorAction SilentlyContinue |
    Where-Object { $_.PNPClass -eq 'Mouse' -and $_.Status -eq 'OK' } |
    Select-Object -First 2)
$mouseStr = if ($mouse.Count) { ($mouse | ForEach-Object { '{0}  [{1}]' -f $_.Name, $_.PNPDeviceID }) -join '  |  ' } else { 'None' }

# Vendor lookup from PNP device ID (VID_xxxx)
$vidMap = @{
    '046D'='Logitech'; '045E'='Microsoft'; '413C'='Dell'; '17EF'='Lenovo'
    '051D'='Lenovo'; '04B3'='Lenovo/IBM'; '8087'='Intel'; '04F2'='Chicony'
    '0461'='Primax'; '09DA'='A4Tech'; '046A'='Cherry'; '1532'='Razer'
    '0C45'='Microdia'; '04D9'='Holtek'; '062A'='MosArt'; '06CB'='Synaptics'
    '04F3'='Elan'; '05AC'='Apple'; '1241'='Belkin'; '0E8F'='GreenAsia'
    '04B4'='Cypress'; '040B'='Weltrend'; '0483'='STMicro'; '04CA'='LiteOn'
    '0489'='LiteOn'; '0BDA'='Realtek'; '0930'='Toshiba'; '1A86'='QinHeng'
    '0A5C'='Broadcom'; '18D1'='Google'; '27C6'='Goodix'
}
function Get-DeviceVendor {
    param([string]$pnpId)
    # USB PNP id: USB\VID_xxxx&PID_yyyy -> look up the vendor
    if ($pnpId -match 'VID_([0-9A-Fa-f]{4})') {
        $vid = $Matches[1].ToUpper()
        if ($vidMap.ContainsKey($vid)) { return $vidMap[$vid] }
        return "VID_$vid"
    }
    # HID / I2C device: HID\INTCxxx, HID\ELANxxx, etc. -> decode the prefix
    if ($pnpId -match 'HID\\([A-Za-z0-9]+)') {
        $prefix = $Matches[1].ToUpper()
        $hidMap = @{
            'INTC'='Intel'; 'ELAN'='Elan'; 'SYNA'='Synaptics'; 'SYN'='Synaptics'
            'WACO'='Wacom'; 'LEN0'='Lenovo'; 'DELL'='Dell'; 'HPQ'='HP'
            'MSFT'='Microsoft'; 'ACPI'='ACPI'; 'SYSTEM'='System (built-in)'
        }
        foreach ($k in $hidMap.Keys) { if ($prefix -like "$k*") { return $hidMap[$k] } }
        return "HID:$prefix"
    }
    return 'Unknown'
}
$keyboardVendor = if ($keyboard.Count) { (Get-DeviceVendor $keyboard[0].PNPDeviceID) } else { 'N/A' }
$mouseVendor    = if ($mouse.Count)    { (Get-DeviceVendor $mouse[0].PNPDeviceID) }    else { 'N/A' }

# Scanner (imaging / WIA / name match)
$scanner = @(Get-CimInstance Win32_PnPEntity -ErrorAction SilentlyContinue |
    Where-Object { $_.Status -eq 'OK' -and ($_.PNPClass -eq 'Image' -or $_.Name -match 'scan|flatbed|scanner') } |
    Select-Object -ExpandProperty Name -Unique)
$scannerStr = if ($scanner.Count) { $scanner -join '  |  ' } else { 'None detected' }

#--------------------------------------------------------------
# Build the row
#--------------------------------------------------------------
$row = [PSCustomObject]@{
    'ComputerName'    = $computerName
    'UserName'        = $userName
    'Owner'           = $ownerName
    'WorkEmail'       = $workEmail
    'Domain'          = $domain
    'OS'              = "$osName $osArch"
    'OSBuild'         = $osBuild
    'OSSupport'       = $osSupport
    'Activation'      = $activation
    'Manufacturer'    = $manufacturer
    'Model'           = $model
    'SerialNumber'    = $serial
    'CPU'             = $cpuName
    'Cores'           = $cores
    'Threads'         = $threads
    'TotalRAM_GB'     = $totalRAM_GB
    'RAM_Modules'     = $ramModules
    'RAM_Speed'       = $ramSpeed
    'RAM_Max_GB'      = $ramMaxGB
    'RAM_FreeSlots'   = $ramFreeSlot
    'Disks'           = $diskSummary
    'DiskHealth'      = $diskHealth
    'DriveC_Total_GB' = $cTotal
    'DriveC_Free_GB'  = $cFree
    'IPv4'            = $ipAddr
    'MAC'             = $macAddr
    'GPU'             = $gpuName
    'Battery'         = $batteryP
    'BatteryHealth_Pct' = $battHealth
    'TPM_Enabled'     = $tpmOn
    'BitLocker'       = $bitlocker
    'Defender'        = $defender
    'Defender_SigAge_Days' = $defSigAge
    'Defender_LastScan_Days' = $defScanAge
    'Firewall'        = $firewall
    'PendingReboot'   = $pendingReboot
    'Monitor(s)'      = $monitors
    'Keyboard'        = $keyboardStr
    'KeyboardVendor'  = $keyboardVendor
    'Mouse'           = $mouseStr
    'MouseVendor'     = $mouseVendor
    'Scanner'         = $scannerStr
    'LastBoot'        = $lastBoot
    'Uptime_Days'     = $uptime
    'Collected'       = (Get-Date -Format 'yyyy-MM-dd HH:mm')
}

#--------------------------------------------------------------
# Save
#--------------------------------------------------------------
if (-not (Test-Path -LiteralPath $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}
# Name the file after the Owner (safe filename), else fall back to computer name
$fileBase = if ($ownerName -and $ownerName.Trim()) { ($ownerName -replace '[^\w\- ]','_' -replace '\s+','_').Trim('_') } else { $computerName }
if (-not $fileBase) { $fileBase = $computerName }
$outFile = Join-Path $OutputFolder ('{0}_sysinfo.csv' -f $fileBase)
$row | Export-Csv -LiteralPath $outFile -NoTypeInformation -Encoding UTF8
Write-Host "[OK] Saved: $outFile" -ForegroundColor Green

if ($Master) {
    $masterFile = Join-Path $OutputFolder 'Master_Inventory.csv'
    $existing = @()
    if (Test-Path -LiteralPath $masterFile) { $existing = @(Import-Csv -LiteralPath $masterFile) }
    $existing = @($existing | Where-Object { $_.ComputerName -ne $computerName })
    $existing += $row
    $existing | Export-Csv -LiteralPath $masterFile -NoTypeInformation -Encoding UTF8
    Write-Host "[OK] Master updated: $masterFile" -ForegroundColor Green
}
