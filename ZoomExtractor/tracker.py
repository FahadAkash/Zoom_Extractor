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

# Win32 imports for window capture
try:
    import win32gui
    import win32ui
    import win32con
    from PIL import Image
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    print("Warning: pywin32 not available. Window capture will use fallback method.")

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
        
        # Window capture support
        self.window_handle = None  # HWND for Win32 capture
        self.window_title = None  # Store window title for tracking
        self.use_window_capture = WIN32_AVAILABLE  # Use Win32 if available

    def capture_window_content(self, hwnd):
        """
        Capture window content using Win32 PrintWindow API
        This works even when window is partially obscured or moved
        
        Args:
            hwnd: Window handle (HWND)
            
        Returns:
            numpy array (BGR image) or None on failure
        """
        if not WIN32_AVAILABLE:
            return None
            
        try:
            # Get window dimensions
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            width = right - left
            height = bottom - top
            
            # Get window DC
            hwndDC = win32gui.GetWindowDC(hwnd)
            mfcDC = win32ui.CreateDCFromHandle(hwndDC)
            saveDC = mfcDC.CreateCompatibleDC()
            
            # Create bitmap
            saveBitMap = win32ui.CreateBitmap()
            saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
            saveDC.SelectObject(saveBitMap)
            
            # PrintWindow captures the window content
            # PW_RENDERFULLCONTENT (0x00000002) for better capture
            result = win32gui.PrintWindow(hwnd, saveDC.GetSafeHdc(), 2)
            
            if result == 0:
                print("PrintWindow failed, trying alternative method...")
                # Try without PW_RENDERFULLCONTENT flag
                result = win32gui.PrintWindow(hwnd, saveDC.GetSafeHdc(), 0)
            
            # Convert to bitmap
            bmpinfo = saveBitMap.GetInfo()
            bmpstr = saveBitMap.GetBitmapBits(True)
            
            # Convert to PIL Image
            img = Image.frombuffer(
                'RGB',
                (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
                bmpstr, 'raw', 'BGRX', 0, 1
            )
            
            # Clean up
            win32gui.DeleteObject(saveBitMap.GetHandle())
            saveDC.DeleteDC()
            mfcDC.DeleteDC()
            win32gui.ReleaseDC(hwnd, hwndDC)
            
            # Convert PIL to numpy array (BGR for OpenCV)
            img_np = np.array(img)
            img_bgr = cv2.cvtColor(img_np, cv2.COLOR_RGB2BGR)
            
            return img_bgr
            
        except Exception as e:
            print(f"Error capturing window content: {e}")
            import traceback
            traceback.print_exc()
            return None

    def find_zoom_window(self):
        """Auto-detect Zoom window"""
        print("Searching for Zoom window...")
        
        # First check if Tesseract is available
        try:
            pytesseract.get_tesseract_version()
        except Exception as e:
            print(f"Warning: Tesseract OCR not available: {e}")
        
        # Patterns to EXCLUDE (our own app)
        exclude_patterns = [
            "zoom attendance system",
            "attendance system",
            "zoom extractor"
        ]
        
        try:
            # Priority 1: Look for actual Zoom meeting windows
            zoom_meeting_patterns = [
                "zoom meeting",
                "zoom webinar",
                "'s zoom meeting",
                "zoom - "
            ]
            
            print("Looking for Zoom meeting windows...")
            for w in gw.getAllTitles():
                if not w.strip():
                    continue
                
                w_lower = w.lower()
                
                # Skip our own app window
                if any(exclude in w_lower for exclude in exclude_patterns):
                    print(f"Skipping own app window: {w}")
                    continue
                
                # Check for meeting patterns
                for pattern in zoom_meeting_patterns:
                    if pattern in w_lower:
                        try:
                            windows = gw.getWindowsWithTitle(w)
                            if not windows:
                                continue
                                
                            win = windows[0]
                            if win.isMinimized:
                                win.restore()
                                time.sleep(0.3)
                            
                            # Validate window size
                            if win.width < 400 or win.height < 300:
                                continue
                            
                            print(f"✓ Found Zoom meeting window: {w}")
                            
                            # Get window handle for Win32 capture
                            if WIN32_AVAILABLE:
                                try:
                                    hwnd = win._hWnd
                                    self.window_handle = hwnd
                                    self.window_title = w
                                    print(f"  Window handle: {hwnd}")
                                except:
                                    print("  Could not get window handle")
                            
                            return {
                                "top": win.top,
                                "left": win.left,
                                "width": win.width,
                                "height": win.height
                            }
                        except Exception as e:
                            print(f"Error accessing window '{w}': {e}")
                            continue
            
            # Priority 2: Look for any window with "zoom" (but not our app)
            print("Looking for any Zoom windows...")
            for w in gw.getAllTitles():
                if not w.strip():
                    continue
                
                w_lower = w.lower()
                
                # Skip our own app
                if any(exclude in w_lower for exclude in exclude_patterns):
                    continue
                
                if "zoom" in w_lower:
                    try:
                        windows = gw.getWindowsWithTitle(w)
                        if not windows:
                            continue
                            
                        win = windows[0]
                        if win.isMinimized:
                            continue
                        
                        if win.width < 400 or win.height < 300:
                            continue
                        
                        print(f"✓ Found Zoom window: {w}")
                        
                        if WIN32_AVAILABLE:
                            try:
                                hwnd = win._hWnd
                                self.window_handle = hwnd
                                self.window_title = w
                                print(f"  Window handle: {hwnd}")
                            except:
                                pass
                        
                        return {
                            "top": win.top,
                            "left": win.left,
                            "width": win.width,
                            "height": win.height
                        }
                    except Exception as e:
                        print(f"Error accessing window '{w}': {e}")
                        continue
                        
        except Exception as e:
            print(f"Error finding window: {e}")
        
        print("✗ Zoom window NOT found.")
        print("Please open a Zoom meeting or manually select the region.")
        return None

    
    def find_zoom_window_by_title(self, title):
        """Find Zoom window by specific title for dynamic tracking"""
        try:
            if title.strip():
                windows = gw.getWindowsWithTitle(title)
                if windows:
                    win = windows[0]
                    if not win.isMinimized:
                        # Update window handle
                        if WIN32_AVAILABLE:
                            try:
                                hwnd = win._hWnd
                                if hwnd != self.window_handle:
                                    self.window_handle = hwnd
                                    print(f"Updated window handle: {hwnd}")
                            except:
                                pass
                        
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
        # DON'T clear window handle - keep it for Win32 capture
        # This allows Win32 capture to work even with manual region selection
        # Overlapping windows won't affect capture if we have window handle
        print(f"Region set: {region}")
        if self.window_handle:
            print(f"  Will use Win32 capture (HWND: {self.window_handle}) + crop to region")
        else:
            print(f"  Will use screen capture (no window handle)")

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
        
        for idx, tile in enumerate(tiles):
            try:
                # Convert to grayscale
                gray = cv2.cvtColor(tile, cv2.COLOR_BGR2GRAY)
                
                # Advanced preprocessing pipeline for better OCR accuracy
                
                # 1. Noise reduction with Gaussian blur
                denoised = cv2.GaussianBlur(gray, (3, 3), 0)
                
                # 2. Contrast enhancement using CLAHE
                clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
                enhanced = clahe.apply(denoised)
                
                # 3. Binarization with Otsu's method
                _, otsu_thresh = cv2.threshold(enhanced, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                
                # 4. Adaptive thresholding for varying lighting
                adaptive_thresh = cv2.adaptiveThreshold(
                    enhanced, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                    cv2.THRESH_BINARY, 11, 2
                )
                
                # Try multiple preprocessing methods and PSM modes
                candidates = []
                
                # Method 1: Otsu threshold with PSM 7 (single line)
                text1 = pytesseract.image_to_string(
                    otsu_thresh,
                    config="--psm 7 --oem 3"
                ).strip()
                if text1:
                    candidates.append(text1)
                
                # Method 2: Adaptive threshold with PSM 7
                text2 = pytesseract.image_to_string(
                    adaptive_thresh,
                    config="--psm 7 --oem 3"
                ).strip()
                if text2:
                    candidates.append(text2)
                
                # Method 3: Enhanced image with PSM 8 (single word)
                text3 = pytesseract.image_to_string(
                    enhanced,
                    config="--psm 8 --oem 3"
                ).strip()
                if text3:
                    candidates.append(text3)
                
                # Method 4: Original grayscale with PSM 6 (uniform block)
                text4 = pytesseract.image_to_string(
                    gray,
                    config="--psm 6 --oem 3"
                ).strip()
                if text4:
                    candidates.append(text4)
                
                # Select best candidate (longest valid text with letters)
                best_text = ""
                for text in candidates:
                    # Clean up text
                    text = ' '.join(text.split())  # Remove extra whitespace
                    
                    # Must contain letters and be reasonable length
                    if text and any(c.isalpha() for c in text) and 2 < len(text) < 60:
                        if len(text) > len(best_text):
                            best_text = text
                
                if best_text:
                    # Apply OCR error corrections
                    corrected_text = self.correct_common_ocr_errors(best_text)
                    
                    # Validate if this looks like a real name
                    if self.is_valid_name(corrected_text):
                        names.append(corrected_text)
                        print(f"Tile {idx+1}: '{best_text}' → '{corrected_text}'")
                    else:
                        print(f"Tile {idx+1}: '{corrected_text}' REJECTED (invalid format)")
                    
            except Exception as e:
                print(f"OCR error on tile {idx+1}: {e}")
                
        return names

    def is_valid_name(self, text):
        """
        Validate if OCR text looks like a valid participant name
        Filters out garbage, UI elements, and junk text
        """
        if not text or len(text) < 2:
            return False
        
        # Must have at least 2 letters
        letter_count = sum(1 for c in text if c.isalpha())
        if letter_count < 2:
            return False
        
        # Reject if too many special characters (indicates garbage)
        special_chars = sum(1 for c in text if not c.isalnum() and c not in ' ()-.')
        if special_chars > 5:
            return False
        
        # Reject if too short or too long
        if len(text) > 60:
            return False
        
        # Reject common UI patterns
        ui_patterns = [
            'zoom attendance',
            'start tracking',
            'stop', 
            'reset data',
            'active:',
            'detected:',
            'matched:',
            'roll number',
            'match %',
            'status',
            'event log',
            'file help',
            '===',
            '___',
            '|||'
        ]
        
        text_lower = text.lower()
        for pattern in ui_patterns:
            if pattern in text_lower:
                return False
        
        # Reject if mostly special characters
        alpha_ratio = letter_count / len(text)
        if alpha_ratio < 0.4:  # At least 40% should be letters
            return False
        
        return True

    def correct_common_ocr_errors(self, text):
        """Correct common OCR errors in names"""
        if not text:
            return ""
            
        # 1. Remove leading/trailing noise characters
        text = text.strip('| [](){}<>.,:;-_')
        
        # 2. Remove specific noise patterns using regex
        import re
        
        # Remove vertical bars inside text (often read from grid lines)
        text = text.replace('|', '')
        
        # Remove (Host, me) and variations - handles missing opening parenthesis too
        # Matches: (Host, me), Host, me), (Host, me, Host, me
        text = re.sub(r'\(?.*?(?:Host|Me|me).*?\)?', '', text, flags=re.IGNORECASE)
        
        # Remove "GBR", "ED" and other common short noise words if they appear isolated
        text = re.sub(r'\b(?:GBR|ED|ER|ft)\b', '', text, flags=re.IGNORECASE)
        
        # Clean up multiple spaces
        text = ' '.join(text.split())
        
        # 3. Apply specific name corrections
        corrections = {
            'Farhad': 'Fahad',
            'Farad': 'Fahad',
            'atash': 'Akash',
            'Akas': 'Akash',
            'Akashhh': 'Akash',
            'Akashh': 'Akash',
            'Fahad Akas': 'Fahad Akash',
        }
        
        for wrong, correct in corrections.items():
            if wrong in text:
                text = text.replace(wrong, correct)
        
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
        # Initialize MSS in this thread (fixes threading error)
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
                    # Restore window handle that was cleared by set_region
                    # (we want to use window tracking for auto-detected windows)
                    if hasattr(self, '_temp_window_handle'):
                        self.window_handle = self._temp_window_handle
                        self.window_title = self._temp_window_title
                    print(f"Auto-detected Zoom window: {region}")
                else:
                    print("Could not auto-detect Zoom window, waiting...")
                    time.sleep(1)
                    continue
            
            # Dynamic window tracking - update region if window has moved
            if self.window_title and self.use_window_capture:
                current_region = self.find_zoom_window_by_title(self.window_title)
                if current_region:
                    if current_region != self.region:
                        print(f"Window moved: {self.region} → {current_region}")
                        self.region = current_region
                else:
                    print(f"Window '{self.window_title}' not found, attempting re-detection...")
                    # Try to re-detect
                    region = self.find_zoom_window()
                    if region:
                        self.region = region
            
            try:
                capture_count += 1
                print(f"\n=== Capture #{capture_count} ===")
                
                screenshot = None
                
                # Try Win32 window capture first if available
                if self.use_window_capture and self.window_handle and WIN32_AVAILABLE:
                    print(f"Using Win32 capture (HWND: {self.window_handle})")
                    window_screenshot = self.capture_window_content(self.window_handle)
                    
                    if window_screenshot is not None:
                        # If we have a manual region set, crop to that region
                        # Convert screen coordinates to window-relative coordinates
                        if self.region:
                            try:
                                # Get window position
                                left, top, right, bottom = win32gui.GetWindowRect(self.window_handle)
                                
                                # Convert region (screen coords) to window-relative coords
                                rel_x = self.region['left'] - left
                                rel_y = self.region['top'] - top
                                rel_w = self.region['width']
                                rel_h = self.region['height']
                                
                                # Crop the captured window to the selected region
                                h, w = window_screenshot.shape[:2]
                                if (rel_x >= 0 and rel_y >= 0 and 
                                    rel_x + rel_w <= w and rel_y + rel_h <= h):
                                    screenshot = window_screenshot[rel_y:rel_y+rel_h, rel_x:rel_x+rel_w]
                                    print(f"  Cropped to region: {rel_w}x{rel_h} at ({rel_x},{rel_y})")
                                else:
                                    # Region outside window, use full window
                                    screenshot = window_screenshot
                                    print(f"  Warning: Region outside window bounds, using full window")
                            except Exception as e:
                                print(f"  Error cropping to region: {e}")
                                screenshot = window_screenshot
                        else:
                            screenshot = window_screenshot
                    else:
                        print("Win32 capture failed, falling back to screen capture...")
                        self.use_window_capture = False  # Disable for this session
                
                # Fallback to screen region capture
                if screenshot is None:
                    print(f"Using screen region capture: {self.region}")
                    screenshot = np.array(self.sct.grab(self.region))
                
                if screenshot is None:
                    print("Failed to capture screenshot")
                    time.sleep(1)
                    continue
                
                print(f"Screenshot captured: {screenshot.shape}")
                
                tiles = self.crop_tiles(screenshot)
                print(f"Created {len(tiles)} tiles")
                
                names = self.extract_names(tiles)
                print(f"Extracted {len(names)} names: {names}")
                
                self.update_participants(names)
                
                print(f"=== Capture #{capture_count} completed ===\n")
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
        print("Stopping tracker...")
        if self.capture_thread and self.capture_thread.is_alive():
            self.capture_thread.join(timeout=1)  # Reduced timeout
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