@echo off
pyinstaller --onefile --noconsole --name DeskPilot main.py
echo.
echo Build complete. Output: dist\DeskPilot.exe
pause
