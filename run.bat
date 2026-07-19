@echo off
chcp 65001 >nul
cd /d "%~dp0"
python -u tray_app.py
if errorlevel 1 pause
