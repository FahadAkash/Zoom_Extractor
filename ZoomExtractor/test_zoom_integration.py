"""
Test script for Zoom meeting integration
"""

import sys
import os

# Add the current directory to the path so we can import the modules
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def test_imports():
    """Test that all required modules can be imported"""
    try:
        from gui import AttendanceApp
        print("✓ GUI module imported successfully")
    except Exception as e:
        print(f"✗ Failed to import GUI module: {e}")
        return False
    
    try:
        from matcher import RollMatcher
        print("✓ Matcher module imported successfully")
    except Exception as e:
        print(f"✗ Failed to import matcher module: {e}")
        return False
    
    # Test selenium imports
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.service import Service
        from webdriver_manager.chrome import ChromeDriverManager
        print("✓ Selenium modules imported successfully")
    except Exception as e:
        print(f"⚠ Selenium modules not available: {e}")
        print("  This is OK if you're not using the Zoom meeting functionality")
    
    # Test faker import
    try:
        from faker import Faker
        print("✓ Faker module imported successfully")
    except Exception as e:
        print(f"⚠ Faker module not available: {e}")
        print("  This is OK if you're not using the Zoom meeting functionality")
    
    return True

def test_zoom_functionality():
    """Test the Zoom meeting functionality"""
    try:
        # This would test the actual Zoom meeting functionality
        # For now, we'll just check that the required modules are available
        from faker import Faker
        from selenium import webdriver
        from selenium.webdriver.chrome.service import Service
        from webdriver_manager.chrome import ChromeDriverManager
        print("✓ All Zoom meeting modules are available")
        return True
    except ImportError as e:
        print(f"⚠ Zoom meeting modules not available: {e}")
        print("  Install selenium, webdriver-manager, and faker to use Zoom meeting functionality")
        return False

if __name__ == "__main__":
    print("Testing Zoom Extractor Integration...")
    print("=" * 50)
    
    if test_imports():
        print("\n✓ All basic imports successful")
    else:
        print("\n✗ Some imports failed")
        sys.exit(1)
    
    print("\nTesting Zoom meeting functionality...")
    test_zoom_functionality()
    
    print("\n" + "=" * 50)
    print("Integration test completed!")