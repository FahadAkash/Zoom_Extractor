"""
Main Entry Point - Zoom Attendance System
"""

import sys
import tkinter as tk
from gui import AttendanceApp


def main():
    """Launch the application"""
    try:
        root = tk.Tk()
        app = AttendanceApp(root)
        root.mainloop()
    except Exception as e:
        print(f"Error launching application: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
