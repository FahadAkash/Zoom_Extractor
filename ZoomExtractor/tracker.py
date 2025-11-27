"""
Zoom Tracker Module
Handles screen capture, OCR, and participant detection
"""

import cv2
import numpy as np
from mss import mss
import pytesseract
import pygetwindow as gw
from pynput import mouse
from datetime import datetime
import threading
import time


class ZoomTracker:
    def __init__(self, callback=None):
        """
        Initialize tracker
        
        Args:
            callback: Function to call with updates (participants_list, event_type)
        """
        self.tile_height = 70
        self.running = False
        self.paused = False
        self.region = None
        self.current_participants = set()
        self.participants_history = {}  # {name: {'joined': time, 'left': time}}
        self.callback = callback
        self.lock = threading.Lock()
        self.capture_thread = None
        self.sct = None
        
    def find_zoom_window(self):
        """Auto-detect Zoom window"""
        print("Searching for Zoom window...")
        try:
            for w in gw.getAllTitles():
                if "zoom" in w.lower() and w.strip():
                    win = gw.getWindowsWithTitle(w)[0]
                    if win.isMinimized:
                        win.restore()
                    print(f"Found Zoom window: {w}")
                    return {
                        "top": win.top,
                        "left": win.left,
                        "width": win.width,
                        "height": win.height
                    }
        except Exception as e:
            print(f"Error finding window: {e}")
        
        print("Zoom window NOT found.")
        return None

    def get_manual_region(self):
        """Manual region selection via mouse drag"""
        print("\nDrag over the Zoom PARTICIPANT LIST area...")
        print("Click and drag with mouse, then release.")
        
        selection_start = None
        selection_end = None
        
        def on_click(x, y, button, pressed):
            nonlocal selection_start, selection_end
            if pressed:
                selection_start = (x, y)
            else:
                selection_end = (x, y)
                return False

        with mouse.Listener(on_click=on_click) as listener:
            listener.join()

        if selection_start and selection_end:
            x1, y1 = selection_start
            x2, y2 = selection_end

            return {
                "top": min(y1, y2),
                "left": min(x1, x2),
                "width": abs(x2 - x1),
                "height": abs(y2 - y1)
            }
        return None

    def set_region(self, region):
        """Set capture region manually"""
        self.region = region

    def set_tile_height(self, height):
        """Update tile height"""
        with self.lock:
            self.tile_height = max(10, height)

    def crop_tiles(self, img):
        """Crop image into participant tiles"""
        tiles = []
        h, w = img.shape[:2]
        th = max(10, self.tile_height)
        num_tiles = h // th

        for i in range(num_tiles):
            y1 = i * th
            y2 = y1 + th
            tiles.append(img[y1:y2, 0:w])

        return tiles

    def extract_names(self, tiles):
        """Extract names from tiles using OCR"""
        names = []
        for tile in tiles:
            try:
                gray = cv2.cvtColor(tile, cv2.COLOR_BGR2GRAY)
                gray = cv2.threshold(gray, 140, 255, cv2.THRESH_BINARY)[1]
                
                text = pytesseract.image_to_string(gray, config="--psm 7").strip()
                
                if 1 < len(text) < 50:  # Filter valid names
                    names.append(text)
            except Exception as e:
                print(f"OCR error: {e}")
                
        return names

    def update_participants(self, detected_names):
        """Update participant list and detect changes"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        new_set = set(detected_names)
        
        with self.lock:
            joined = new_set - self.current_participants
            left = self.current_participants - new_set
            
            if joined:
                for name in joined:
                    self.participants_history[name] = {
                        'joined': timestamp,
                        'left': None
                    }
                if self.callback:
                    self.callback(list(new_set), 'joined', list(joined))
            
            if left:
                for name in left:
                    if name in self.participants_history:
                        self.participants_history[name]['left'] = timestamp
                if self.callback:
                    self.callback(list(new_set), 'left', list(left))

            self.current_participants = new_set

    def get_attendance_data(self):
        """Get complete attendance data"""
        with self.lock:
            return {
                'current': list(self.current_participants),
                'history': dict(self.participants_history)
            }

    def capture_loop(self):
        """Main capture loop (runs in thread)"""
        self.sct = mss()
        
        while self.running:
            if self.paused or not self.region:
                time.sleep(0.1)
                continue
                
            try:
                screenshot = np.array(self.sct.grab(self.region))
                tiles = self.crop_tiles(screenshot)
                names = self.extract_names(tiles)
                self.update_participants(names)
                
                time.sleep(0.5)  # Adjust capture frequency
                
            except Exception as e:
                print(f"Capture error: {e}")
                time.sleep(1)

    def start(self):
        """Start tracking"""
        if not self.region:
            raise ValueError("Region not set. Call find_zoom_window() or get_manual_region() first.")
            
        self.running = True
        self.capture_thread = threading.Thread(target=self.capture_loop, daemon=True)
        self.capture_thread.start()
        print("Tracker started")

    def stop(self):
        """Stop tracking"""
        self.running = False
        if self.capture_thread:
            self.capture_thread.join(timeout=2)
        print("Tracker stopped")

    def pause(self):
        """Pause tracking"""
        self.paused = True

    def resume(self):
        """Resume tracking"""
        self.paused = False

    def reset(self):
        """Reset all data"""
        with self.lock:
            self.current_participants = set()
            self.participants_history = {}
