@echo off
rem Report maker installer. KEEP THIS FILE ASCII-ONLY (cmd mis-parses multi-byte batch files).
rem Korean messages are printed by install.ps1.
title report maker setup
set "PS1=%TEMP%\report-maker-setup.ps1"
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest ('https://raw.githubusercontent.com/pagecraft-tools/report-maker/main/install.ps1?t=' + [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()) -OutFile '%PS1%' -UseBasicParsing"
if not exist "%PS1%" goto :failed
powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%"
if errorlevel 1 goto :failed
del "%PS1%" >nul 2>&1
pause
exit /b 0
:failed
echo.
echo   Install failed - check your internet connection and try again.
echo.
pause
exit /b 1
