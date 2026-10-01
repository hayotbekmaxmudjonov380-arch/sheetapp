@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo ============================================
echo  SheetApp .exe yig'ish (boshqa foydalanuvchi)
echo ============================================
echo.
echo 1-qadam: pyinstaller o'rnatilganmi tekshirilmoqda...
python -c "import PyInstaller" 2>nul || pip install pyinstaller
echo.
echo 2-qadam: .exe yig'ilmoqda (bir necha daqiqa)...
python -m PyInstaller --noconfirm --onefile --windowed ^
  --name "SheetApp" ^
  --collect-all PySide6 ^
  --hidden-import sheets_app ^
  run_sheetapp.py
echo.
if exist "dist\SheetApp.exe" (
    echo TAYYOR: dist\SheetApp.exe
    echo Bu faylni boshqa kompyuterga nusxalash kifoya —
    echo uchun Python o'rnatish shart emas.
) else (
    echo Yig'ishda xato yuz berdi.
)
pause
