@echo off
cd /d "%~dp0"
echo === start ===
python refresh.py
echo.
start "" "%~dp0안전재고발주.html"
pause
