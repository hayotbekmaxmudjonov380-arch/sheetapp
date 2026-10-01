@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo SheetApp ishga tushmoqda...
python -m sheets_app
if errorlevel 1 (
    echo.
    echo XATOLIK: Python yoki PySide6 topilmadi.
    echo Quyidagini bajaring:  pip install PySide6 openpyxl
    pause
)
