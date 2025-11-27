#!/bin/bash
# Build script for Linux/Mac

echo "==================================="
echo "Zoom Attendance System - Builder"
echo "==================================="
echo ""

# Check if virtual environment exists
if [ ! -d "venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
fi

# Activate virtual environment
echo "Activating virtual environment..."
source venv/bin/activate

# Install dependencies
echo "Installing dependencies..."
pip install -r requirements.txt

# Build with PyInstaller
echo ""
echo "Building executable..."
pyinstaller --onefile \
    --windowed \
    --name="ZoomAttendance" \
    --add-data="tracker.py:." \
    --add-data="matcher.py:." \
    --add-data="gui.py:." \
    main.py

echo ""
echo "==================================="
echo "Build complete!"
echo "Executable location: dist/ZoomAttendance"
echo "==================================="
echo ""
echo "Note: Make sure Tesseract OCR is installed on the system"
echo "Linux: sudo apt install tesseract-ocr"
echo "Mac: brew install tesseract"
