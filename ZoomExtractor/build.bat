@echo off
REM Build script for Windows

echo ===================================
echo Zoom Attendance System - Builder
echo ===================================
echo.

REM Check if virtual environment exists
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
)

REM Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate.bat

REM Install dependencies
echo Installing dependencies...
pip install -r requirements.txt

REM Build with PyInstaller
echo.
echo Building executable...
pyinstaller --onefile ^
    --windowed ^
    --name="ZoomAttendance" ^
    --add-data="tracker.py;." ^
    --add-data="matcher.py;." ^
    --add-data="gui.py;." ^
    main.py

echo.
echo ===================================
echo Build complete!
echo Executable location: dist\ZoomAttendance.exe
echo ===================================
echo.
echo Note: Make sure Tesseract OCR is installed on the system
echo Download from: https://github.com/UB-Mannheim/tesseract/wiki
echo.
pause
