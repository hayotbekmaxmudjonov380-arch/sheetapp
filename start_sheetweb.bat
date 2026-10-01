@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo SheetApp WEB ishga tushmoqda... (http://127.0.0.1:5001)
start "" http://127.0.0.1:5001
python -m sheets_app.web
if errorlevel 1 (
    echo.
    echo XATOLIK: Python yoki Flask topilmadi.
    echo O'rnatish:  pip install Flask openpyxl
    pause
)
