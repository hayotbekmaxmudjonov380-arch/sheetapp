@echo off
chcp 65001 > nul
echo Eskilarni tozalash...
for /f "tokens=5" %%a in ('netstat -a -n -o ^| findstr :5000') do taskkill /f /pid %%a > nul 2>&1
echo Restoran hisob dasturi ishga tushmoqda...
start http://127.0.0.1:5000
python app_web.py
pause
