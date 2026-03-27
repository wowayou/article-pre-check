@echo off
chcp 65001 >nul
set "SCRIPT_PATH=%~dp0SEO_Packer_Pro.ps1"
PowerShell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_PATH%"
pause