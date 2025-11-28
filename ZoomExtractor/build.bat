@echo off
REM Build script for Windows - Standalone Executable

echo ===================================
echo Zoom Attendance System - Standalone Builder
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

REM Upgrade pip
echo Upgrading pip...
python -m pip install --upgrade pip

REM Install dependencies
echo Installing dependencies...
pip install -r requirements.txt

REM Install PyInstaller if not already installed
echo Installing PyInstaller...
pip install pyinstaller

REM Build with PyInstaller - Enhanced configuration for standalone executable
echo.
echo Building standalone executable...
pyinstaller --onefile ^
    --windowed ^
    --name="ZoomAttendance" ^
    --icon=NONE ^
    --hidden-import="selenium" ^
    --hidden-import="webdriver_manager" ^
    --hidden-import="faker" ^
    --hidden-import="pandas" ^
    --hidden-import="openpyxl" ^
    --hidden-import="rapidfuzz" ^
    --hidden-import="pywin32" ^
    --hidden-import="PIL" ^
    --hidden-import="requests" ^
    --hidden-import="pyperclip" ^
    --hidden-import="tkinter" ^
    --hidden-import="tkinter.ttk" ^
    --hidden-import="tkinter.filedialog" ^
    --hidden-import="tkinter.messagebox" ^
    --hidden-import="tkinter.scrolledtext" ^
    --add-data="tracker.py;." ^
    --add-data="matcher.py;." ^
    --add-data="gui.py;." ^
    --add-data="zoommeeting.py;." ^
    --collect-all="selenium" ^
    --collect-all="webdriver_manager" ^
    --collect-all="faker" ^
    --collect-all="pandas" ^
    --collect-all="rapidfuzz" ^
    --collect-all="PIL" ^
    --collect-all="requests" ^
    --collect-all="pyperclip" ^
    main.py

echo.
echo ===================================
echo Build complete!
echo Executable location: dist\ZoomAttendance.exe
echo ===================================
echo.
echo To run the application, double-click on ZoomAttendance.exe in the dist folder
echo This standalone executable includes all necessary dependencies
echo.
pause
