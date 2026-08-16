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

echo ************************************************************
echo  IT Inventory Sync  ^(SharePoint: %LIST_NAME%^)
echo  Reads CSVs from Results\ and updates the list.
echo  A browser window opens once - sign in with work account.
echo ************************************************************
echo.

if defined CLIENT_ID (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Sync-ToSharePoint.ps1" -SiteUrl "%SITE_URL%" -ListName "%LIST_NAME%" -ClientId "%CLIENT_ID%"
) else (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Sync-ToSharePoint.ps1" -SiteUrl "%SITE_URL%" -ListName "%LIST_NAME%"
)

echo.
pause
