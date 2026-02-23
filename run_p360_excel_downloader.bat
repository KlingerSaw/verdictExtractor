@echo off
setlocal
chcp 65001 > nul

set "SCRIPT_DIR=%~dp0"
set "PS1_FILE=%SCRIPT_DIR%Download-P360Files-Excel.ps1"

if not exist "%PS1_FILE%" (
    echo [FEJL] Fandt ikke scriptfilen:
    echo        "%PS1_FILE%"
    echo.
    echo Loesning: Kopier .bat og Download-P360Files-Excel.ps1 til samme mappe
    echo eller opdater din genvej, saa den peger paa den korrekte placering.
    pause
    exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS1_FILE%"
pause
