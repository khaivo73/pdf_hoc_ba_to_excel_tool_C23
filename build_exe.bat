@echo off
chcp 65001 >nul
cd /d "%~dp0"

python -m pip install --upgrade pip
if errorlevel 1 goto fail

python -m pip install -r requirements.txt
if errorlevel 1 goto fail

python -m PyInstaller ^
  --noconfirm ^
  --onefile ^
  --windowed ^
  --name HocBaC23PdfToExcel ^
  --add-data "templates;templates" ^
  --exclude-module PyQt6 ^
  --exclude-module PySide6 ^
  --exclude-module PyQt5 ^
  --exclude-module PySide2 ^
  --exclude-module torch ^
  --exclude-module pandas ^
  --exclude-module scipy ^
  --exclude-module IPython ^
  --exclude-module matplotlib ^
  app.py
if errorlevel 1 goto fail

echo.
echo DONE. File EXE nam tai: dist\HocBaC23PdfToExcel.exe
pause
exit /b 0

:fail
echo.
echo BUILD FAILED. Vui long xem loi o tren.
pause
exit /b 1
