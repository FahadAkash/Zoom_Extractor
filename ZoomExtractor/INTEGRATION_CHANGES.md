# Zoom Meeting Integration Changes

## Overview
This document describes the changes made to integrate the zoommeeting.py functionality into the main GUI application.

## Files Modified

### 1. gui.py
Major changes to replace the screen capture approach with the Selenium-based Zoom meeting approach:

#### Removed Functionality:
- `auto_detect_region()` method
- `manual_select_region()` method
- `update_tile_height()` method
- `on_tracker_update()` method
- `refresh_zoom_window()` method (functionality changed)

#### Added Functionality:
- Zoom meeting input fields (Meeting ID, Passcode, Participants)
- `_join_zoom_meeting()` method - Joins Zoom meetings programmatically
- `_fetch_participants()` method - Extracts participant names from Zoom interface
- `_monitor_participants()` method - Monitors participant data and updates UI
- Updated `start_tracking()` method to use Zoom meeting approach
- Updated `stop_tracking()` method to handle meeting termination
- Updated `refresh_zoom_window()` to show informational message

#### Updated Dependencies:
- Added imports for selenium, faker, and related modules
- Added queue module for participant data handling
- Added threading event for graceful shutdown

### 2. requirements.txt
Added new dependencies required for Selenium-based approach:
- selenium
- webdriver-manager
- faker

## Key Improvements

### 1. Direct Participant Extraction
Instead of using OCR on screen captures, the new approach:
- Joins Zoom meetings programmatically
- Extracts participant names directly from the Zoom interface
- Provides more accurate and reliable participant detection

### 2. Configurable Participant Count
Users can now specify how many participants to join the meeting with, allowing for:
- Testing with multiple virtual participants
- Simulating different class sizes
- Better attendance tracking

### 3. Improved User Interface
- Simplified setup with direct input fields
- Removed complex region selection process
- More intuitive workflow

### 4. Better Error Handling
- Graceful handling of meeting termination
- Improved logging and status updates
- Better error messages for missing dependencies

## Backward Compatibility
The screen capture approach has been removed from the GUI but the underlying code remains in tracker.py for reference or future use.

## Testing
A test script (`test_zoom_integration.py`) has been added to verify:
- Module imports
- Basic functionality
- Dependency availability

## Usage Instructions

### New Workflow:
1. Load roll number database
2. Enter Zoom meeting details:
   - Meeting ID
   - Passcode
   - Number of participants
3. Click "Join Meeting"
4. Monitor participant detection in real-time
5. Export attendance data when needed

### Dependencies:
Users must install additional dependencies:
```bash
pip install selenium webdriver-manager faker
```

If these dependencies are not available, the system will show a warning but continue to function for other features.