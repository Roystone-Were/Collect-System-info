@echo off
title IT Inventory - Sync to SharePoint
cd /d "%~dp0"

REM ============================================================
REM  Your SharePoint site and inventory list:
set "SITE_URL=https://refrontiergroup.sharepoint.com/sites/xanalifeTechData"
set "LIST_NAME=System Information"
REM  If Register-PnPEntraIDApp gave you a ClientId, remove the
REM  'REM' below and paste your guid in:
REM set "CLIENT_ID=paste-guid-here"
REM ============================================================

REM  PnP.PowerShell needs PowerShell 7+ (pwsh); fall back to 5.1 only
REM  so the script can print its own "needs PowerShell 7" message.
set "PS_EXE=pwsh.exe"
where pwsh.exe >nul 2>nul || set "PS_EXE=powershell.exe"

echo ************************************************************
echo  IT Inventory Sync  ^(SharePoint: %LIST_NAME%^)
echo  Using: %PS_EXE%
echo  Reads CSVs from Results\ and updates the list.
echo  A browser window opens once - sign in with work account.
echo ************************************************************
echo.

if defined CLIENT_ID (
    "%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Sync-ToSharePoint.ps1" -SiteUrl "%SITE_URL%" -ListName "%LIST_NAME%" -ClientId "%CLIENT_ID%"
) else (
    "%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Sync-ToSharePoint.ps1" -SiteUrl "%SITE_URL%" -ListName "%LIST_NAME%"
)

echo.
pause
