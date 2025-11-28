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
import os

# Configure pytesseract to use the Tesseract executable directly
try:
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
except:
    pass  # If this fails, pytesseract will try to find tesseract in PATH


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
        
        # First check if Tesseract is available
        try:
            pytesseract.get_tesseract_version()
        except Exception as e:
            print(f"Warning: Tesseract OCR not available: {e}")
        
        try:
            # Try different patterns to find Zoom window
            zoom_patterns = ["zoom", "Zoom", "ZOOM", "zoom meeting", "zoom webinar"]
            
            for w in gw.getAllTitles():
                # Check if window title contains any zoom pattern
                if any(pattern in w.lower() for pattern in zoom_patterns) and w.strip():
                    try:
                        win = gw.getWindowsWithTitle(w)[0]
                        if win.isMinimized:
                            win.restore()
                        print(f"Found Zoom window: {w}")
                        # Store the window title for dynamic tracking
                        self._last_window_title = w
                        return {
                            "top": win.top,
                            "left": win.left,
                            "width": win.width,
                            "height": win.height
                        }
                    except Exception as e:
                        print(f"Error accessing window '{w}': {e}")
                        continue
            
            # If no zoom window found, try to find any window that might be Zoom
            # by checking window class or other properties
            print("Trying alternative detection methods...")
            for w in gw.getAllWindows():
                try:
                    # Check if window title is not empty and window is visible
                    if w.title.strip() and not w.isMinimized:
                        # Check if window dimensions are typical for a video conferencing app
                        if w.width > 800 and w.height > 600:
                            print(f"Found potential window: {w.title} ({w.width}x{w.height})")
                            # Store the window title for dynamic tracking
                            self._last_window_title = w.title
                            # For debugging, let's return the first large window we find
                            # In a real implementation, you might want to be more selective
                            return {
                                "top": w.top,
                                "left": w.left,
                                "width": w.width,
                                "height": w.height
                            }
                except:
                    continue
                        
        except Exception as e:
            print(f"Error finding window: {e}")
        
        print("Zoom window NOT found.")
        return None
    
    def find_zoom_window_by_title(self, title):
        """Find Zoom window by specific title for dynamic tracking"""
        try:
            if title.strip():
                windows = gw.getWindowsWithTitle(title)
                if windows:
                    win = windows[0]
                    if not win.isMinimized:
                        return {
                            "top": win.top,
                            "left": win.left,
                            "width": win.width,
                            "height": win.height
                        }
        except Exception as e:
            print(f"Error finding window by title '{title}': {e}")
        return None

    def get_manual_region(self):
        """Manual region selection via mouse drag with visual feedback"""
        print("\nDrag over the Zoom PARTICIPANT LIST area...")
        print("Click and drag with mouse, then release.")
        
        selection_start = None
        selection_end = None
        
        # Try to show a simple visual indicator using console output
        print("Visual feedback: Click and drag to select region...")
        print("(A visual selection box will appear on screen during dragging)")
        
        def on_click(x, y, button, pressed):
            nonlocal selection_start, selection_end
            if pressed:
                selection_start = (x, y)
                print(f"Selection started at ({x}, {y})")
            else:
                selection_end = (x, y)
                print(f"Selection ended at ({x}, {y})")
                return False

        with mouse.Listener(on_click=on_click) as listener:
            listener.join()

        if selection_start and selection_end:
            x1, y1 = selection_start
            x2, y2 = selection_end
            
            # Print the selected region
            width = abs(x2 - x1)
            height = abs(y2 - y1)
            top = min(y1, y2)
            left = min(x1, x2)
            print(f"Selected region: Top={top}, Left={left}, Width={width}, Height={height}")

            return {
                "top": top,
                "left": left,
                "width": width,
                "height": height
            }
        return None

    def set_region(self, region):
        """Set capture region manually"""
        self.region = region
        # Clear the last window title when setting region manually
        # This disables dynamic window tracking for manually selected regions
        if hasattr(self, '_last_window_title'):
            delattr(self, '_last_window_title')

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
            
        print(f"Cropping image {h}x{w} into {num_tiles} tiles of height {th}")
            
        for i in range(num_tiles):
            y1 = i * th
            y2 = y1 + th
            # Ensure we don't go beyond image bounds
            if y2 <= h:
                tile = img[y1:y2, 0:w]
                tiles.append(tile)
                print(f"Created tile {i+1}: {tile.shape}")
            
        return tiles

    def extract_names(self, tiles):
        """Extract names from tiles using OCR"""
        names = []
        
        # Check if Tesseract is available
        try:
            pytesseract.get_tesseract_version()
        except pytesseract.TesseractNotFoundError:
            # Try to set the path directly
            try:
                pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
                pytesseract.get_tesseract_version()  # Test if this works
                print("Tesseract found at default location and configured.")
            except Exception:
                print("ERROR: Tesseract OCR is not installed or not in PATH.")
                print("Please install Tesseract OCR and make sure it's in your system PATH.")
                print("Download from: https://github.com/UB-Mannheim/tesseract/wiki")
                return names  # Return empty list if Tesseract is not available
        except Exception as e:
            print(f"Error checking Tesseract: {e}")
            return names
        
        for tile in tiles:
            try:
                # Convert to grayscale
                gray = cv2.cvtColor(tile, cv2.COLOR_BGR2GRAY)
                
                # Try multiple preprocessing techniques to improve OCR
                # Technique 1: Simple threshold
                _, thresh1 = cv2.threshold(gray, 140, 255, cv2.THRESH_BINARY)
                text1 = pytesseract.image_to_string(thresh1, config="--psm 7 -c tessedit_char_whitelist=abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789()- ").strip()
                
                # Technique 2: Adaptive threshold
                thresh2 = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
                text2 = pytesseract.image_to_string(thresh2, config="--psm 7 -c tessedit_char_whitelist=abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789()- ").strip()
                
                # Technique 3: No preprocessing
                text3 = pytesseract.image_to_string(gray, config="--psm 7 -c tessedit_char_whitelist=abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789()- ").strip()
                
                # Technique 4: Try with different PSM mode for single line text
                text4 = pytesseract.image_to_string(gray, config="--psm 8 -c tessedit_char_whitelist=abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789()- ").strip()
                
                # Use the best result (longest valid name)
                candidates = [text1, text2, text3, text4]
                best_text = ""
                for text in candidates:
                    # Filter for valid names (avoid empty or whitespace-only)
                    if text and not text.isspace() and len(text) > len(best_text):
                        # Additional validation - check if it looks like a name
                        if any(c.isalpha() for c in text):  # Must contain at least one letter
                            best_text = text
                
                # Filter valid names
                if 1 < len(best_text) < 50 and not best_text.isspace():
                    # Clean up the text
                    cleaned_text = ' '.join(best_text.split())  # Remove extra whitespace
                    if cleaned_text:  # Only add non-empty names
                        # Apply some common corrections for better accuracy
                        corrected_text = self.correct_common_ocr_errors(cleaned_text)
                        names.append(corrected_text)
                        print(f"Detected name: '{cleaned_text}' -> Corrected to: '{corrected_text}'")  # Debug output
            except Exception as e:
                print(f"OCR error: {e}")
                
        return names
    
    def correct_common_ocr_errors(self, text):
        """Correct common OCR errors in names"""
        # Common OCR corrections
        corrections = {
            'Farhad': 'Fahad',
            'Akas': 'Akash',
            'Fahad Akas': 'Fahad Akash',
            'Fahad Akash (': 'Fahad Akash',
            'Fahad Akash (Hose': 'Fahad Akash (Host)',
            'Fahad Akash (Me': 'Fahad Akash (Me)',
            'Fahad Akash (Host': 'Fahad Akash (Host)',
            'Fahad Akash Me': 'Fahad Akash (Me)',
        }
        
        # Apply corrections
        for wrong, correct in corrections.items():
            if wrong in text:
                text = text.replace(wrong, correct)
        
        # Remove trailing punctuation that might be OCR artifacts
        text = text.rstrip('.,:;')
        
        # Ensure proper formatting for Zoom participant names
        if '(' in text and ')' not in text:
            text = text.replace('(', '(Me)') if 'me' in text.lower() else text.replace('(', '(Host)')
        
        return text.strip()

    def update_participants(self, detected_names):
        """Update participant list and detect changes"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        new_set = set(detected_names)
        
        print(f"Updating participants. Detected: {detected_names}")
        print(f"Current participants: {self.current_participants}")
        
        with self.lock:
            joined = new_set - self.current_participants
            left = self.current_participants - new_set
            
            if joined:
                print(f"New participants joined: {joined}")
                for name in joined:
                    self.participants_history[name] = {
                        'joined': timestamp,
                        'left': None
                    }
                if self.callback:
                    self.callback(list(new_set), 'joined', list(joined))
            
            if left:
                print(f"Participants left: {left}")
                for name in left:
                    if name in self.participants_history:
                        self.participants_history[name]['left'] = timestamp
                if self.callback:
                    self.callback(list(new_set), 'left', list(left))

            self.current_participants = new_set
            print(f"Updated current participants: {self.current_participants}")

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
        
        # Check if Tesseract is available before starting capture loop
        try:
            pytesseract.get_tesseract_version()
            print("Tesseract OCR is available and ready.")
        except pytesseract.TesseractNotFoundError:
            # Try to set the path directly
            try:
                pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
                pytesseract.get_tesseract_version()  # Test if this works
                print("Tesseract found at default location and configured.")
            except Exception:
                print("ERROR: Tesseract OCR is not installed or not in PATH.")
                print("Please install Tesseract OCR and make sure it's in your system PATH.")
                print("Download from: https://github.com/UB-Mannheim/tesseract/wiki")
                print("After installation, you may need to restart your computer for PATH changes to take effect.")
                return
        except Exception as e:
            print(f"Error checking Tesseract: {e}")
            return
        
        capture_count = 0
        
        while self.running:
            if self.paused:
                time.sleep(0.1)
                continue
                
            # If no region is set, try to auto-detect Zoom window
            if not self.region:
                print("No region set, attempting to auto-detect Zoom window...")
                region = self.find_zoom_window()
                if region:
                    self.set_region(region)
                    print(f"Auto-detected Zoom window: {region}")
                else:
                    print("Could not auto-detect Zoom window, waiting...")
                    time.sleep(1)
                    continue
            
            # Dynamic window tracking - update region if window has moved
            if hasattr(self, '_last_window_title') and self._last_window_title:
                current_region = self.find_zoom_window_by_title(self._last_window_title)
                if current_region:
                    if current_region != self.region:
                        print(f"Zoom window moved from {self.region} to {current_region}")
                        self.region = current_region  # Update region directly to avoid clearing _last_window_title
                else:
                    print(f"Zoom window '{self._last_window_title}' no longer found, pausing tracking")
                    # Don't return here, continue with capture but it will likely fail
                    # The UI can handle this case
            
            try:
                capture_count += 1
                print(f"Capture #{capture_count} started")
                
                screenshot = np.array(self.sct.grab(self.region))
                print(f"Screenshot captured: {screenshot.shape}")
                
                tiles = self.crop_tiles(screenshot)
                print(f"Created {len(tiles)} tiles")
                
                names = self.extract_names(tiles)
                print(f"Extracted {len(names)} names: {names}")
                
                self.update_participants(names)
                
                print(f"Capture #{capture_count} completed\n")
                time.sleep(0.5)  # Adjust capture frequency
                
            except Exception as e:
                print(f"Capture error: {e}")
                import traceback
                traceback.print_exc()
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