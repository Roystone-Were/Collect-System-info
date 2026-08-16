@echo off
title IT Inventory - Sync to SharePoint
cd /d "%~dp0"

REM ============================================================
REM  EDIT THIS LINE ONCE - your SharePoint site address:
set "SITE_URL=https://YOURCOMPANY.sharepoint.com/sites/YOURSITE"
REM  If Register-PnPEntraIDApp gave you a ClientId, remove the
REM  'REM' below and paste your guid in:
REM set "CLIENT_ID=paste-guid-here"
REM ============================================================

if "%SITE_URL%"=="https://YOURCOMPANY.sharepoint.com/sites/YOURSITE" (
    echo ************************************************************
    echo  First time: right-click this file - Edit, and set
    echo  SITE_URL to your SharePoint site address.
    echo ************************************************************
    pause
    exit /b
)

echo ************************************************************
echo  IT Inventory Sync  (SharePoint list)
echo  Reads CSVs from Results\ and updates the list.
echo  A browser window opens once - sign in with work account.
echo ************************************************************
echo.

if defined CLIENT_ID (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Sync-ToSharePoint.ps1" -SiteUrl "%SITE_URL%" -ClientId "%CLIENT_ID%"
) else (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Sync-ToSharePoint.ps1" -SiteUrl "%SITE_URL%"
)

echo.
pause
