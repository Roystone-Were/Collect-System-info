# Collect System Info

USB-based IT fleet inventory collector for Windows 10/11. Plug the USB into a
laptop, double-click `Run-on-USB.cmd`, and a hardware/user inventory CSV is
saved straight to the USB stick - no network share or admin rights needed.

## Usage

1. Copy this folder to a USB stick.
2. On each laptop, double-click `Run-on-USB.cmd` (run as the logged-in user).
3. It runs `Get-ITInventory.ps1` and writes `<ComputerName>_sysinfo.csv`
   into `Results\` on the USB.
4. Bring the USB back and merge all CSVs into one workbook with
   `Merge-ITInventory.ps1` (kept in the central ITInventory folder).

## What each CSV captures

Computer name · logged-in user · owner · work email · domain · Windows
edition/build · manufacturer · model · serial number · CPU (cores/threads) ·
RAM (GB/modules/speed) · disks · C: total/free · IPv4 · MAC · GPU · battery ·
TPM · monitors · keyboard/mouse/scanner · last boot · uptime · timestamp.

## Notes

- No admin rights required (reads WMI/CIM + the user's registry hive).
- Run as the logged-in user for best work-email detection.
- `Results/` is git-ignored on purpose: collected CSVs contain PII
  (usernames, emails, serials, MACs). Keep them on restricted storage only.
