# Zoom Attendance System

A production-ready attendance tracking system for Zoom meetings with automatic name-to-roll-number matching using OCR.

## Features

✨ **Automatic Attendance Tracking** - Captures participant names from Zoom using screen OCR  
🎯 **Roll Number Matching** - Fuzzy matching to link detected names with roll numbers  
🖥️ **Full GUI** - User-friendly interface with real-time tracking  
📊 **Multiple Export Formats** - Excel, CSV reports  
📦 **Standalone Executable** - Build as `.exe` for Windows or binary for Linux/Mac  
🔔 **Event Detection** - Real-time join/leave notifications  

---

## Installation

### Prerequisites
- Python 3.8+
- Tesseract OCR installed on your system

**Install Tesseract:**
- **Linux**: `sudo apt install tesseract-ocr`
- **Mac**: `brew install tesseract`
- **Windows**: Download from [here](https://github.com/UB-Mannheim/tesseract/wiki)

### Setup

1. **Clone/Download** this repository

2. **Create virtual environment:**
   ```bash
   python3 -m venv venv
   ```

3. **Activate virtual environment:**
   - Linux/Mac: `source venv/bin/activate`
   - Windows: `venv\Scripts\activate`

4. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

---

## Usage

### Running the Application

```bash
python main.py
```

### Step-by-Step Guide

1. **Setup Tab**
   - Click **"Browse File"** to upload your roll number file (`.txt` format)
   - Select capture region:
     - **Auto-Detect**: Automatically finds Zoom window
     - **Manual**: Click and drag to select participant list area
   - Adjust settings (tile height, match threshold)

2. **Live Tracking Tab**
   - Click **"▶ Start Tracking"** to begin
   - View real-time participant list with matched roll numbers
   - Monitor join/leave events in the log
   - Click **"⏹ Stop"** to end session

3. **Reports Tab**
   - View session summary
   - Export attendance to Excel or CSV

---

## Roll Number File Format

Create a text file with one entry per line:

```
Fahad Akash 08
John Doe 15
Sarah Smith 22
```

**Format:** `Name RollNumber` (name followed by roll number)

---

## Building Standalone Executable

### Linux/Mac

```bash
chmod +x build.sh
./build.sh
```

### Windows

```batch
build.bat
```

The executable will be in the `dist/` folder.

---

## Configuration

### Settings (Setup Tab)

- **Tile Height**: Adjust to match the height of one participant row in Zoom (typically 60-80 pixels)
- **Match Threshold**: Minimum similarity percentage for name matching (75% recommended)

---

## Troubleshooting

### "No module named 'cv2'" error
- Make sure you're in the virtual environment
- Run: `pip install -r requirements.txt`

### Tesseract not found
- Install Tesseract OCR on your system
- Linux: `sudo apt install tesseract-ocr`

### Names not detected
- Adjust **Tile Height** in settings
- Ensure participant list is clearly visible
- Check region selection covers the participant names

### Low match accuracy
- Lower the **Match Threshold** in settings
- Ensure roll file names match Zoom display names closely

---

## File Structure

```
zoom/
├── main.py              # Application entry point
├── gui.py               # GUI interface
├── tracker.py           # Screen capture & OCR logic
├── matcher.py           # Name-to-roll matching
├── requirements.txt     # Python dependencies
├── build.sh             # Linux/Mac build script
├── build.bat            # Windows build script
└── README.md            # This file
```

---

## Technologies Used

- **Python 3** - Core language
- **Tkinter** - GUI framework
- **OpenCV** - Image processing
- **Tesseract OCR** - Text recognition
- **RapidFuzz** - Fuzzy string matching
- **Pandas** - Data handling
- **PyInstaller** - Executable builder

---

## License

Free to use and modify.

---

## Support

For issues or questions, refer to the troubleshooting section above.
