@echo off
setlocal

rem Finds HELP+.accdb automatically. To use a specific file instead,
rem drag the .accdb onto this .bat file.
set "NAME=HELP+.accdb"

set "OUT=%~dp0html"
set "PS64=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "PS32=%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
set "ARGS=-NoProfile -ExecutionPolicy Bypass -File "%~dp0export-to-html.ps1" -OutDir "%OUT%" -DbName "%NAME%""
if not "%~1"=="" set "ARGS=%ARGS% -DbPath "%~1""

rem Clear the old index so a failed run doesn't open a stale snapshot
if exist "%OUT%\index.html" del "%OUT%\index.html"

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
