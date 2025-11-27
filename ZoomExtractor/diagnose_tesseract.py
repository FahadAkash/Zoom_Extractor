"""
Diagnostic script to check Tesseract OCR installation
"""

import sys
import os
import subprocess

def check_tesseract_command():
    """Check if tesseract command is available"""
    try:
        result = subprocess.run(['tesseract', '--version'], 
                              capture_output=True, text=True, timeout=10)
        if result.returncode == 0:
            print("✓ Tesseract command is available")
            print(f"Version: {result.stdout.splitlines()[0]}")
            return True
        else:
            print("✗ Tesseract command failed")
            print(f"Error: {result.stderr}")
            return False
    except FileNotFoundError:
        print("✗ Tesseract command not found in PATH")
        return False
    except Exception as e:
        print(f"✗ Error checking tesseract command: {e}")
        return False

def check_python_tesseract():
    """Check if pytesseract Python package can find Tesseract"""
    try:
        import pytesseract
        
        # Try to get version directly
        version = pytesseract.get_tesseract_version()
        print("✓ Python pytesseract package can find Tesseract")
        print(f"Version: {version}")
        return True
    except ImportError:
        print("✗ Python pytesseract package not installed")
        return False
    except pytesseract.TesseractNotFoundError:
        print("✗ Python pytesseract cannot find Tesseract executable")
        # Try to set the path directly
        try:
            pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
            version = pytesseract.get_tesseract_version()
            print("✓ Tesseract found at default location")
            print(f"Version: {version}")
            return True
        except Exception as e2:
            print(f"✗ Tesseract not found at default location either: {e2}")
            return False
    except Exception as e:
        print(f"✗ Error with pytesseract: {e}")
        return False

def check_path():
    """Check PATH for common Tesseract locations"""
    print("\nChecking PATH for Tesseract...")
    path_dirs = os.environ.get('PATH', '').split(os.pathsep)
    
    # Common Tesseract installation paths on Windows
    common_paths = [
        r'C:\Program Files\Tesseract-OCR',
        r'C:\Program Files (x86)\Tesseract-OCR',
        r'C:\Tesseract-OCR'
    ]
    
    found_paths = []
    for path_dir in path_dirs:
        if 'tesseract' in path_dir.lower():
            found_paths.append(path_dir)
    
    for common_path in common_paths:
        if os.path.exists(common_path) and common_path not in path_dirs:
            print(f"⚠ Tesseract found at '{common_path}' but not in PATH")
    
    if found_paths:
        print("✓ Tesseract paths found in PATH:")
        for path in found_paths:
            print(f"  - {path}")
    else:
        print("✗ No Tesseract paths found in PATH")

def main():
    print("=== Tesseract OCR Diagnostic Tool ===\n")
    
    # Check Python version
    print(f"Python version: {sys.version}")
    
    # Check tesseract command
    print("\n1. Checking tesseract command...")
    cmd_ok = check_tesseract_command()
    
    # Check Python pytesseract
    print("\n2. Checking Python pytesseract...")
    python_ok = check_python_tesseract()
    
    # Check PATH
    print("\n3. Checking PATH...")
    check_path()
    
    # Summary
    print("\n=== Summary ===")
    if cmd_ok and python_ok:
        print("✓ Tesseract is properly installed and configured!")
        print("You should be able to use the Zoom Extractor application.")
    else:
        print("✗ Tesseract is not properly configured.")
        print("\nTo fix this issue:")
        print("1. Download Tesseract from: https://github.com/UB-Mannheim/tesseract/wiki")
        print("2. Install Tesseract (use default installation path)")
        print("3. Add Tesseract to your PATH environment variable")
        print("4. Restart your computer")
        print("5. Run this diagnostic tool again to verify")

if __name__ == "__main__":
    main()