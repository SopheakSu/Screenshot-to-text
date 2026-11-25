@echo off
title Build ScreenToExcel.exe
cd /d "%~dp0"
where python >nul 2>nul || (echo Python not found. Install Python 3.10+ from python.org and tick "Add python.exe to PATH". & pause & exit /b 1)
echo Installing required packages...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt || (echo Package install failed. & pause & exit /b 1)
echo Building exe...
python -m PyInstaller --noconfirm --clean --onefile --windowed --name ScreenToExcel ^
  --collect-all winrt ^
  --hidden-import winrt.windows.media.ocr ^
  --hidden-import winrt.windows.globalization ^
  --hidden-import winrt.windows.graphics.imaging ^
  --hidden-import winrt.windows.storage.streams ^
  --hidden-import winrt.windows.foundation ^
  --hidden-import winrt.windows.foundation.collections ^
  ScreenToExcel.py || (echo Build failed. & pause & exit /b 1)
echo.
echo Done!  Your app is:  %~dp0dist\ScreenToExcel.exe
explorer "%~dp0dist"
pause
