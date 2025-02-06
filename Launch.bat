@echo off
set "scriptPath=%~dp0AzureManagement.ps1"
PowerShell -NoProfile -ExecutionPolicy Bypass -File "%scriptPath%"
pause
