@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo [1/3] Install deps...
python -m pip install -q -r requirements.txt pyinstaller

echo [2/3] Build onefile exe...
python -m PyInstaller --noconfirm --clean ^
  --onefile ^
  --windowed ^
  --name "TBH通关时间监控" ^
  --add-data "clear_time_probe.js;." ^
  --hidden-import excel_store ^
  --hidden-import frida ^
  --hidden-import pystray._win32 ^
  --collect-all frida ^
  --collect-all pystray ^
  --collect-all openpyxl ^
  tray_app.py

if errorlevel 1 (
  echo BUILD FAILED
  exit /b 1
)

echo [3/3] Done.
echo EXE: %~dp0dist\TBH通关时间监控.exe
dir /b "dist"
