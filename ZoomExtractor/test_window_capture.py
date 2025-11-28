"""
Test Win32 Window Capture - Quick Verification
Tests if pywin32 and window capture functionality works correctly
"""

import sys

def test_imports():
    """Test if all required imports are available"""
    print("=" * 60)
    print("Testing Required Imports")
    print("=" * 60)
    
    try:
        import win32gui
        import win32ui
        import win32con
        print("✓ pywin32 modules imported successfully")
    except ImportError as e:
        print(f"✗ Failed to import pywin32: {e}")
        print("\nPlease install: pip install pywin32")
        return False
    
    try:
        from PIL import Image
        print("✓ PIL (Pillow) imported successfully")
    except ImportError as e:
        print(f"✗ Failed to import PIL: {e}")
        print("\nPlease install: pip install pillow")
        return False
    
    try:
        import cv2
        print("✓ OpenCV imported successfully")
    except ImportError as e:
        print(f"✗ Failed to import OpenCV: {e}")
        return False
    
    try:
        import pytesseract
        print("✓ pytesseract imported successfully")
    except ImportError as e:
        print(f"✗ Failed to import pytesseract: {e}")
        return False
    
    return True

def test_window_detection():
    """Test window detection functionality"""
    print("\n" + "=" * 60)
    print("Testing Window Detection")
    print("=" * 60)
    
    try:
        import pygetwindow as gw
        
        windows = gw.getAllTitles()
        visible_windows = [w for w in windows if w.strip()]
        
        print(f"\nFound {len(visible_windows)} visible windows:")
        for i, title in enumerate(visible_windows[:10], 1):
            print(f"  {i}. {title[:60]}...")
        
        if len(visible_windows) > 10:
            print(f"  ... and {len(visible_windows) - 10} more")
        
        return True
        
    except Exception as e:
        print(f"✗ Window detection failed: {e}")
        return False

def test_win32_capture():
    """Test Win32 capture capability"""
    print("\n" + "=" * 60)
    print("Testing Win32 Capture Capability")
    print("=" * 60)
    
    try:
        import win32gui
        import pygetwindow as gw
        
        # Try to find any window to test capture
        windows = [w for w in gw.getAllTitles() if w.strip()]
        
        if not windows:
            print("✗ No windows found to test")
            return False
        
        # Get first available window
        test_window_title = windows[0]
        wins = gw.getWindowsWithTitle(test_window_title)
        
        if wins:
            win = wins[0]
            hwnd = win._hWnd
            print(f"✓ Successfully got window handle: {hwnd}")
            print(f"  Window: '{test_window_title[:50]}...'")
            
            # Get window rect to verify access
            rect = win32gui.GetWindowRect(hwnd)
            print(f"✓ Window dimensions: {rect[2]-rect[0]}x{rect[3]-rect[1]}")
            
            return True
        else:
            print("✗ Could not access window")
            return False
            
    except Exception as e:
        print(f"✗ Win32 capture test failed: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_tracker_import():
    """Test if tracker module can be imported with Win32 support"""
    print("\n" + "=" * 60)
    print("Testing Tracker Module")
    print("=" * 60)
    
    try:
        from tracker import ZoomTracker, WIN32_AVAILABLE
        
        if WIN32_AVAILABLE:
            print("✓ Tracker imported with Win32 support enabled")
        else:
            print("⚠ Tracker imported but Win32 support disabled")
            return False
        
        # Create tracker instance
        tracker = ZoomTracker()
        print("✓ ZoomTracker instance created successfully")
        
        # Check for new methods
        if hasattr(tracker, 'capture_window_content'):
            print("✓ capture_window_content() method available")
        else:
            print("✗ capture_window_content() method missing")
            return False
        
        # Check for new attributes
        if hasattr(tracker, 'window_handle'):
            print("✓ window_handle attribute available")
        else:
            print("✗ window_handle attribute missing")
            return False
        
        if hasattr(tracker, 'use_window_capture'):
            print(f"✓ use_window_capture = {tracker.use_window_capture}")
        else:
            print("✗ use_window_capture attribute missing")
            return False
        
        return True
        
    except Exception as e:
        print(f"✗ Failed to import tracker: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """Run all tests"""
    print("\n")
    print("╔" + "=" * 58 + "╗")
    print("║" + " " * 10 + "Win32 Window Capture - Verification Test" + " " * 7 + "║")
    print("╚" + "=" * 58 + "╝")
    print()
    
    results = []
    
    # Run tests
    results.append(("Import Test", test_imports()))
    if results[-1][1]:  # Only continue if imports work
        results.append(("Window Detection", test_window_detection()))
        results.append(("Win32 Capture", test_win32_capture()))
        results.append(("Tracker Module", test_tracker_import()))
    
    # Print summary
    print("\n" + "=" * 60)
    print("Test Summary")
    print("=" * 60)
    
    for test_name, passed in results:
        status = "✓ PASS" if passed else "✗ FAIL"
        print(f"{status:8} | {test_name}")
    
    print("=" * 60)
    
    all_passed = all(result[1] for result in results)
    
    if all_passed:
        print("\n🎉 All tests passed! Window capture system is ready.")
        print("\nYou can now run the main application:")
        print("  python main.py")
    else:
        print("\n⚠️  Some tests failed. Please resolve the issues above.")
        print("\nCommon fixes:")
        print("  1. Install dependencies: pip install -r requirements.txt")
        print("  2. Restart your terminal/IDE")
        print("  3. Check if pywin32 installed correctly")
    
    print()
    return 0 if all_passed else 1

if __name__ == "__main__":
    sys.exit(main())
