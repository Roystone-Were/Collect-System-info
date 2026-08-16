@echo off
title One-time setup - Entra App Registration (IT Inventory Sync)
cd /d "%~dp0"

echo ************************************************************
echo  ONE-TIME SETUP
echo  A browser window opens: sign in with your ADMIN account
echo  and click consent/accept. Then a ClientId (a guid) is
echo  printed below - copy it, it goes into Run-Sync.cmd.
echo ************************************************************
echo.

where pwsh.exe >nul 2>nul || (
    echo PowerShell 7 ^(pwsh^) not found. Install it first ^(Microsoft Store: "PowerShell"^).
    pause
    exit /b
)

pwsh -NoProfile -Command "Register-PnPEntraIDApp -Interactive -DisplayName 'IT Inventory Sync' -Tenant 'refrontiergroup.onmicrosoft.com'"

echo.
echo ************************************************************
echo  Copy the ClientId shown above, then:
echo  right-click Run-Sync.cmd - Edit - remove the REM on the
echo  CLIENT_ID line and paste the guid there.
echo ************************************************************
pause
