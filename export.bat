@echo off
setlocal

rem ==== Change this to your database path ====
set "DB=\\server\share\HELP+.accdb"
rem ===========================================

set "OUT=%~dp0html"
set "PS64=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "PS32=%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
set "ARGS=-NoProfile -ExecutionPolicy Bypass -File "%~dp0export-to-html.ps1" -DbPath "%DB%" -OutDir "%OUT%""

"%PS64%" %ARGS%
rem Exit code 2 = driver not found in 64-bit; 32-bit Office needs 32-bit PowerShell
if errorlevel 2 if exist "%PS32%" "%PS32%" %ARGS%

if exist "%OUT%\index.html" (
    start "" "%OUT%\index.html"
) else (
    echo.
    echo Export failed - see messages above.
    pause
)
