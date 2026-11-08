@echo off
cd /d "%~dp0"
powershell.exe -NoLogo -NoProfile -STA -File "%~dp0Install.ps1"
if errorlevel 1 pause
