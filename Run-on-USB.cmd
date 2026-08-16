@echo off
title IT Inventory - collect to this USB
cd /d "%~dp0"
echo ************************************************************
echo  IT Inventory Collector  (USB MODE)
echo  Saves the CSV directly onto this USB stick.
echo ************************************************************
echo.
echo Collecting system info...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Get-ITInventory.ps1" -OutputFolder "%~dp0Results" -AskOwner
echo.
echo Done. Saved to: %~dp0Results
echo Bring this USB back to the IT lead to merge.
pause
