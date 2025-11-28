# Zoom Attendance System

## Overview
The Zoom Attendance System automatically tracks Zoom meeting attendance and matches participants with roll numbers from a database. This system provides two approaches for participant detection:

1. **Screen Capture Method** (Legacy): Uses OCR to detect participants from screen captures
2. **Selenium Method** (New): Joins Zoom meetings programmatically and extracts participant information directly

## New Selenium-Based Approach

The new approach integrates the zoommeeting.py functionality directly into the GUI:

### Features
- Automatically join Zoom meetings with configurable number of participants
- Extract participant names directly from the Zoom interface
- Match participants with roll numbers using fuzzy matching
- Real-time participant tracking and logging

### Requirements
- Python 3.7+
- Chrome browser
- ChromeDriver (automatically managed by webdriver-manager)
- Selenium
- Faker

### Installation
```bash
pip install -r requirements.txt
```

### Usage
1. Launch the application:
   ```bash
   python main.py
   ```

2. In the Setup tab:
   - Load your roll number database file
   - Enter Zoom meeting ID
   - Enter meeting passcode
   - Set number of participants to join (1-100)

3. Switch to the Live Tracking tab and click "Join Meeting"

4. The system will:
   - Join the Zoom meeting with the specified number of participants
   - Extract participant names from the meeting interface
   - Match names with roll numbers
   - Display results in real-time

## Legacy Screen Capture Method

The legacy method is still available in the codebase but disabled in the GUI.

## Configuration

### Roll Number Database
Create a text file with student names and roll numbers in one of these formats:
```
Fahad Akash 08
Jahid Hasan 15
1. Fahad Akash
2. Jahid Hasan
```

### Settings
- Match Threshold: Adjust the fuzzy matching sensitivity (50-100%)

## Exporting Data
- Export to Excel (.xlsx)
- Export to CSV (.csv)

## Troubleshooting

### Missing Dependencies
If you see warnings about missing selenium or faker modules:
```bash
pip install selenium webdriver-manager faker
```

### ChromeDriver Issues
The system uses webdriver-manager to automatically download and manage ChromeDriver.
If you encounter issues, ensure Chrome is installed and up to date.

### Participant Detection Issues
If participants are not being detected correctly:
1. Ensure the Zoom meeting interface is visible
2. Check that the participant panel is open
3. Adjust the match threshold in settings