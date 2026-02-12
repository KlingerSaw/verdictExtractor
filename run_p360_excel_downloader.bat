@echo off
chcp 65001 > nul
powershell.exe -ExecutionPolicy Bypass -File "%~dp0Download-P360Files-Excel.ps1"
pause
