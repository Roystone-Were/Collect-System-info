# Collect System Info

USB-based IT fleet inventory collector for Windows 10/11. Plug the USB into a
laptop, double-click `Run-on-USB.cmd`, and a hardware/user inventory CSV is
saved straight to the USB stick - no network share or admin rights needed.

## Usage

1. Copy this folder to a USB stick.
2. On each laptop, double-click `Run-on-USB.cmd` (run as the logged-in user).
3. It runs `Get-ITInventory.ps1` (asking you to type the PC owner's name)
   and writes `<Owner>_sysinfo.csv` into `Results\` on the USB.
4. Back at your desk, combine every CSV into one workbook + master CSV:
   `powershell -NoProfile -ExecutionPolicy Bypass -File .\Merge-ITInventory.ps1 -Folder .\Results`
5. (Optional/experimental) push everything into a SharePoint list:
   `powershell -NoProfile -ExecutionPolicy Bypass -File .\Sync-ToSharePoint.ps1 -SiteUrl "https://<tenant>.sharepoint.com/sites/<site>"`
   Requires PowerShell 7 and a one-time Entra app registration - see the
   script header. The Excel master from step 4 is the primary output.

## What each CSV captures

Computer name · logged-in user · owner · work email · domain · Windows
edition/build + support status + activation · manufacturer · model · serial
number · CPU (cores/threads) · RAM (GB/modules/speed/max/empty slots) · disks
+ health status · C: total/free · IPv4 · MAC · GPU · battery + health % ·
TPM · BitLocker · Defender (AV/real-time/signature age/last scan) · firewall
per profile · pending reboot · monitors · keyboard/mouse/scanner · last
boot · uptime · timestamp.

Note: BitLocker needs the collector to run elevated (right-click,
Run as administrator); all other fields work as a normal user.

## Notes

- No admin rights required (reads WMI/CIM + the user's registry hive).
- Run as the logged-in user for best work-email detection.
- `Results/` is git-ignored on purpose: collected CSVs contain PII
  (usernames, emails, serials, MACs). Keep them on restricted storage only.
