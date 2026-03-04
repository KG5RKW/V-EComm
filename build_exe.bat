@echo off
echo ==========================================
echo Building EXE...
echo ==========================================

pyinstaller --noconfirm --clean ^
 --onefile ^
 --windowed ^
 --name V-EComm ^
 --icon V-EComm.ico ^
 V-EComm.py

echo.
echo BUILD COMPLETE!
pause