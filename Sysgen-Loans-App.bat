@echo off
title Sysgen RPO - Loans Intelligence Suite
cd /d "%~dp0"

echo ========================================================
echo   SYSGEN RPO - LOANS INTELLIGENCE SUITE
echo   Authorized Personnel: Ivy Salamanca
echo   Engineered by: Christian Antero Salamanca
echo ========================================================
echo.
echo Launching Desktop Application...
echo.

:: Check if dev server is responding on port 5173
powershell -Command "$client = New-Object System.Net.Sockets.TcpClient; try { $client.Connect('127.0.0.1', 5173); exit 0 } catch { exit 1 }" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Starting local application engine...
    start /min "Sysgen Loans Engine" npm run dev
    timeout /t 2 /nobreak >nul
)

:: Launch Edge in Native Standalone App Window Mode
where msedge >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    start msedge --app=http://localhost:5173/ --window-size=1440,900
    exit /b
)

:: Fallback to Chrome in Native App Window Mode
where chrome >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    start chrome --app=http://localhost:5173/ --window-size=1440,900
    exit /b
)

:: Fallback: Open default browser
start http://localhost:5173/
exit /b
