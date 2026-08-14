@echo off
chcp 65001 >nul
title HE THONG NHAP LIEU BIEN CHE 2026
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo [1/3] Dang tao moi truong Python...
  python -m venv .venv
  if errorlevel 1 goto :python_error
)
echo [2/3] Dang cai thu vien can thiet...
call .venv\Scripts\activate.bat
python -m pip install -r requirements.txt
if errorlevel 1 goto :pip_error
echo [3/3] Dang khoi dong he thong...
echo.
echo Mo trinh duyet tai: http://127.0.0.1:5000
echo Tai khoan Admin ban dau: admin
echo Mat khau tam: Admin@2026!
echo.
python app.py
pause
exit /b
:python_error
echo KHONG TIM THAY PYTHON. Hay cai Python 3.11/3.12 va tick "Add Python to PATH".
pause
exit /b 1
:pip_error
echo Khong cai duoc thu vien. Kiem tra ket noi Internet roi chay lai file nay.
pause
exit /b 1
