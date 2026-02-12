@echo off
chcp 65001 > nul
powershell.exe -ExecutionPolicy Bypass -File "%~dp0Download-P360Files-SIF.ps1"
pause
