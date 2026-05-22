@echo off
REM Launcher cho HocBaPdfToExcel - kiểm tra và cập nhật tự động
REM 
REM Cách sử dụng:
REM   run_launcher.bat              (chạy ứng dụng desktop với kiểm tra update)
REM   run_launcher.bat --skip-update (bỏ qua kiểm tra update)
REM   run_launcher.bat --web        (chạy ứng dụng web)

cd /d "%~dp0"

REM Cài đặt dependencies nếu chưa có
echo.
echo Checking dependencies...
python -m pip install -q requests

REM Chạy launcher
python launcher.py %*

pause
