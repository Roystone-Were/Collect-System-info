@echo off
title IT Inventory - Merge all laptops into Excel
cd /d "%~dp0"
echo ************************************************************
echo  IT Inventory MERGE
echo  Combines all CSVs in Results\ into Master_Inventory.xlsx
echo  (one row per machine, newest wins)
echo ************************************************************
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Merge-ITInventory.ps1" -Folder "%~dp0Results"
echo.
echo Master files are in the Results folder.
echo Press any key to close.
pause
