"""
GUI Module - Zoom Attendance System
Main application interface using Tkinter
"""

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import threading
from datetime import datetime
import pandas as pd
import requests
from matcher import RollMatcher
import time
import queue
import pyperclip

# Import zoommeeting functionality
try:
    from faker import Faker
    from selenium import webdriver
    from selenium.webdriver.support.wait import WebDriverWait
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as ec
    from selenium.webdriver.chrome.service import Service
    from webdriver_manager.chrome import ChromeDriverManager
    ZOOM_MEETING_AVAILABLE = True
except ImportError:
    ZOOM_MEETING_AVAILABLE = False
    print("Warning: Zoom meeting functionality not available. Install selenium and faker.")


class AttendanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Zoom Attendance System - Automated Attendance Tracker")
        self.root.geometry("1000x750")
        self.root.minsize(900, 700)
        
        # Set modern theme colors
        # Theme is handled by ttkbootstrap in main()
        pass
        
        # Initialize components
        self.matcher = RollMatcher()
        
        # State
        self.is_tracking = False
        self.roll_file_loaded = False
        self.meeting_active = False
        self.continuous_save_enabled = False  # Continuous save feature toggle
        self.participants_queue = queue.Queue()  # Queue for participant data
        self.meeting_thread = None
        self.stop_event = threading.Event()
        
        # Refresh counter for participant data fetching
        self.refresh_count = 0
        
        # Create UI
        self.create_menu()
        self.create_tabs()
        self.create_status_bar()
        
    def create_menu(self):
        """Create menu bar"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Load Roll Numbers", command=self.load_roll_file, accelerator="Ctrl+O")
        file_menu.add_command(label="Load from Google Sheet", command=self.load_google_sheet)
        file_menu.add_separator()
        file_menu.add_command(label="Export to Excel", command=self.export_excel, accelerator="Ctrl+E")
        file_menu.add_command(label="Export to CSV", command=self.export_csv, accelerator="Ctrl+Shift+E")
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit, accelerator="Ctrl+Q")
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="User Guide", command=self.show_help)
        help_menu.add_command(label="About", command=self.show_about)
        
        # Bind keyboard shortcuts
        self.root.bind('<Control-o>', lambda e: self.load_roll_file())
        self.root.bind('<Control-e>', lambda e: self.export_excel())
        self.root.bind('<Control-E>', lambda e: self.export_excel())
        self.root.bind('<Control-q>', lambda e: self.root.quit())
        self.root.bind('<Control-Q>', lambda e: self.root.quit())
        
    def create_tabs(self):
        """Create tabbed interface"""
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Setup Tab
        self.setup_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.setup_tab, text="Setup")
        self.create_setup_tab()
        
        # Live Tab
        self.live_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.live_tab, text="Live Tracking")
        self.create_live_tab()
        
        # Report Tab
        self.report_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.report_tab, text="Reports")
        self.create_report_tab()
        
    def create_setup_tab(self):
        """Setup tab - configuration"""
        # Add welcome message
        welcome_frame = ttk.Frame(self.setup_tab, padding=10)
        welcome_frame.pack(fill=tk.X, padx=10, pady=5)
        
        welcome_label = ttk.Label(welcome_frame, 
                                 text="Welcome to Zoom Attendance System!\nConfigure your settings below to start tracking attendance.", 
                                 font=('Arial', 10), 
                                 justify=tk.CENTER)
        welcome_label.pack()
        
        # Roll number file section
        frame_roll = ttk.Labelframe(self.setup_tab, text="Student Database", padding=15)
        frame_roll.pack(fill=tk.X, padx=15, pady=10)
        
        ttk.Label(frame_roll, text="Load student data with format: 'Name Roll'", font=('Arial', 9, 'bold')).pack(anchor=tk.W)
        ttk.Label(frame_roll, text="Example: Fahad Akash 08", font=('Arial', 9, 'italic')).pack(anchor=tk.W)
        
        btn_frame = ttk.Frame(frame_roll)
        btn_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(btn_frame, text="📁 Browse File", command=self.load_roll_file, bootstyle=PRIMARY).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🔗 From Google Sheet", command=self.load_google_sheet, bootstyle=INFO).pack(side=tk.LEFT, padx=5)
        self.roll_status = ttk.Label(btn_frame, text="No file loaded", foreground="red", font=('Arial', 9))
        self.roll_status.pack(side=tk.LEFT, padx=10)
        
        # Course code section
        frame_course = ttk.Labelframe(self.setup_tab, text="Course Information", padding=15)
        frame_course.pack(fill=tk.X, padx=15, pady=10)
        
        course_frame = ttk.Frame(frame_course)
        course_frame.pack(fill=tk.X)
        
        ttk.Label(course_frame, text="Course Code:", font=('Arial', 9)).pack(side=tk.LEFT)
        self.course_code_var = tk.StringVar(value="CSE - 407")
        course_entry = ttk.Entry(course_frame, textvariable=self.course_code_var, width=20)
        course_entry.pack(side=tk.LEFT, padx=10)
        
        # Zoom Meeting Details
        frame_meeting = ttk.Labelframe(self.setup_tab, text="Zoom Meeting Configuration", padding=15)
        frame_meeting.pack(fill=tk.X, padx=15, pady=10)
        
        # Meeting ID
        ttk.Label(frame_meeting, text="Meeting ID:", font=('Arial', 9)).grid(row=0, column=0, sticky=tk.W, pady=8)
        self.meeting_id_var = tk.StringVar()
        ttk.Entry(frame_meeting, textvariable=self.meeting_id_var, width=25, font=('Arial', 10)).grid(row=0, column=1, sticky=tk.W, padx=5)
        
        # Passcode
        ttk.Label(frame_meeting, text="Passcode:", font=('Arial', 9)).grid(row=1, column=0, sticky=tk.W, pady=8)
        self.passcode_var = tk.StringVar()
        ttk.Entry(frame_meeting, textvariable=self.passcode_var, width=25, show="●", font=('Arial', 10)).grid(row=1, column=1, sticky=tk.W, padx=5)
        
        # Number of Participants
        ttk.Label(frame_meeting, text="Participants:", font=('Arial', 9)).grid(row=2, column=0, sticky=tk.W, pady=8)
        self.participants_var = tk.IntVar(value=1)
        ttk.Spinbox(frame_meeting, from_=1, to=100, textvariable=self.participants_var, width=10, font=('Arial', 10)).grid(row=2, column=1, sticky=tk.W, padx=5)
        
        # Settings
        frame_settings = ttk.Labelframe(self.setup_tab, text="Advanced Settings", padding=15)
        frame_settings.pack(fill=tk.X, padx=15, pady=10)
        
        ttk.Label(frame_settings, text="Match Threshold (%):", font=('Arial', 9)).grid(row=0, column=0, sticky=tk.W, pady=8)
        self.threshold_var = tk.IntVar(value=self.matcher.threshold)
        ttk.Scale(frame_settings, from_=50, to=100, variable=self.threshold_var, 
                 orient=tk.HORIZONTAL, command=self.update_threshold, length=200).grid(row=0, column=1, sticky=tk.EW, padx=5)
        self.threshold_label = ttk.Label(frame_settings, text=str(self.matcher.threshold), font=('Arial', 9, 'bold'))
        self.threshold_label.grid(row=0, column=2, padx=10)
        
        # Continuous Save Option
        self.continuous_save_var = tk.BooleanVar(value=False)
        continuous_save_check = ttk.Checkbutton(frame_settings, 
                                               text="Enable Continuous Save (Auto-save report periodically)", 
                                               variable=self.continuous_save_var, 
                                               command=self.toggle_continuous_save,
                                               bootstyle="round-toggle")
        continuous_save_check.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        # Add help text
        help_label = ttk.Label(frame_settings, 
                              text="Tip: Higher threshold = stricter matching\n"
                              "Continuous save protects against data loss",
                              font=('Arial', 8), 
                              foreground='gray')
        help_label.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        frame_settings.columnconfigure(1, weight=1)
        
    def create_live_tab(self):
        """Live tracking tab"""
        # Header
        header_frame = ttk.Frame(self.live_tab, padding=10)
        header_frame.pack(fill=tk.X, padx=10, pady=5)
        
        header_label = ttk.Label(header_frame, 
                                text="Live Attendance Tracking\nMonitor participants in real-time during Zoom meetings", 
                                font=('Arial', 11), 
                                justify=tk.CENTER)
        header_label.pack()
        
        # Controls
        control_frame = ttk.Labelframe(self.live_tab, text="Meeting Controls", padding=15)
        control_frame.pack(fill=tk.X, padx=15, pady=10)
        
        self.btn_start = ttk.Button(control_frame, text="▶ Join Meeting", 
                                     command=self.start_tracking, bootstyle=SUCCESS)
        self.btn_start.pack(side=tk.LEFT, padx=10)
        
        self.btn_stop = ttk.Button(control_frame, text="⏹ Stop Tracking", 
                                    command=self.stop_tracking, state=tk.DISABLED, bootstyle=DANGER)
        self.btn_stop.pack(side=tk.LEFT, padx=10)
        
        self.btn_reset = ttk.Button(control_frame, text="↺ Reset Data", command=self.reset_data, bootstyle=WARNING)
        self.btn_reset.pack(side=tk.LEFT, padx=10)
        
        # Stats
        stats_frame = ttk.Labelframe(self.live_tab, text="Attendance Statistics", padding=15)
        stats_frame.pack(fill=tk.X, padx=15, pady=10)
        
        self.stat_active = ttk.Label(stats_frame, text="Active: 0", font=('Arial', 14, 'bold'), foreground='#3498db')
        self.stat_active.pack(side=tk.LEFT, padx=25)
        
        self.stat_total = ttk.Label(stats_frame, text="Total Detected: 0", font=('Arial', 14, 'bold'), foreground='#f39c12')
        self.stat_total.pack(side=tk.LEFT, padx=25)
        
        self.stat_matched = ttk.Label(stats_frame, text="Matched: 0", font=('Arial', 14, 'bold'), foreground='#27ae60')
        self.stat_matched.pack(side=tk.LEFT, padx=25)
        
        # Refresh counter indicator
        self.stat_refresh = ttk.Label(stats_frame, text="Refreshes: 0", font=('Arial', 14, 'bold'), foreground='#9b59b6')
        self.stat_refresh.pack(side=tk.LEFT, padx=25)
        
        # Participant list
        list_frame = ttk.Labelframe(self.live_tab, text="Participant List", padding=10)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
        
        # Treeview
        columns = ('Name', 'Roll', 'Confidence', 'Status')
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=12)
        
        self.tree.heading('Name', text='Participant Name')
        self.tree.heading('Roll', text='Roll Number')
        self.tree.heading('Confidence', text='Match %')
        self.tree.heading('Status', text='Match Status')
        
        self.tree.column('Name', width=300)
        self.tree.column('Roll', width=120)
        self.tree.column('Confidence', width=100)
        self.tree.column('Status', width=120)
        
        # Configure treeview tags for styling
        self.tree.tag_configure('matched', foreground='#27ae60', background='#e8f5e9', font=('Arial', 9, 'bold'))
        self.tree.tag_configure('unmatched', foreground='#7f8c8d', background='#f8f9fa')
        
        scrollbar_v = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar_h = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscroll=scrollbar_v.set, xscroll=scrollbar_h.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_v.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_h.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Event log
        log_frame = ttk.Labelframe(self.live_tab, text="System Events", padding=10)
        log_frame.pack(fill=tk.X, padx=15, pady=10)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=6, state=tk.DISABLED, font=('Consolas', 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
    def create_report_tab(self):
        """Reports tab"""
        # Header
        header_frame = ttk.Frame(self.report_tab, padding=10)
        header_frame.pack(fill=tk.X, padx=10, pady=5)
        
        header_label = ttk.Label(header_frame, 
                                text="Attendance Report\nView, export, or copy attendance data", 
                                font=('Arial', 12), 
                                justify=tk.CENTER)
        header_label.pack()
        
        frame = ttk.Labelframe(self.report_tab, text="Attendance Summary", padding=15)
        frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
        
        ttk.Label(frame, text="Session Report", font=('Arial', 16, 'bold')).pack(pady=10)
        
        self.report_text = scrolledtext.ScrolledText(frame, height=20, font=('Consolas', 10))
        self.report_text.pack(fill=tk.BOTH, expand=True, pady=15)
        
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(btn_frame, text="🔄 Refresh Report", command=self.generate_report, bootstyle=PRIMARY).pack(side=tk.LEFT, padx=8)
        ttk.Button(btn_frame, text="📊 Export Excel", command=self.export_excel, bootstyle=SUCCESS).pack(side=tk.LEFT, padx=8)
        ttk.Button(btn_frame, text="📑 Export CSV", command=self.export_csv, bootstyle=INFO).pack(side=tk.LEFT, padx=8)
        ttk.Button(btn_frame, text="📋 Copy to Clipboard", command=self.copy_attendance_to_clipboard, bootstyle=SECONDARY).pack(side=tk.LEFT, padx=8)
        
    def create_status_bar(self):
        """Status bar at bottom"""
        self.status_bar = ttk.Label(self.root, 
                                  text="Ready - Zoom Attendance System v1.0", 
                                  relief=tk.SUNKEN, 
                                  anchor=tk.W, 
                                  padding=5,
                                  background='#ecf0f1',
                                  foreground='#2c3e50',
                                  font=('Arial', 9))
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
    # Event handlers
    def load_roll_file(self):
        """Load roll number file"""
        filepath = filedialog.askopenfilename(
            title="Select Roll Number File",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
        )
        
        if filepath:
            try:
                count = self.matcher.load_from_file(filepath)
                self.roll_file_loaded = True
                self.roll_status.config(text=f"✓ {count} records loaded", foreground="green")
                self.log(f"Loaded {count} roll numbers from file", "success")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file:\n{e}")
                
    def load_google_sheet(self):
        """Load roll numbers from Google Sheet"""
        # Create a dialog to get the Google Sheet URL
        dialog = tk.Toplevel(self.root)
        dialog.title("Google Sheet URL")
        dialog.geometry("500x150")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        ttk.Label(dialog, text="Enter Google Sheets URL:").pack(pady=10)
        
        # Use default URL
        default_url = "https://docs.google.com/spreadsheets/d/11_2OwNaav5TrzuteEAekweMMYopRbhMhpuCRxvtXzhU/edit?gid=0#gid=0"
        url_var = tk.StringVar(value=default_url)
        entry = ttk.Entry(dialog, textvariable=url_var, width=60)
        entry.pack(padx=10, pady=5)
        entry.focus()
        
        def load_sheet():
            url = url_var.get().strip()
            if url:
                try:
                    count = self.matcher.load_from_google_sheet(url)
                    self.roll_file_loaded = True
                    self.roll_status.config(text=f"✓ {count} records loaded", foreground="green")
                    self.log(f"Loaded {count} roll numbers from Google Sheet", "success")
                    dialog.destroy()
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to load Google Sheet:\n{e}")
            else:
                messagebox.showwarning("Warning", "Please enter a valid URL")
        
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="Load", command=load_sheet, bootstyle=PRIMARY).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy, bootstyle=SECONDARY).pack(side=tk.LEFT, padx=5)
        
        # Bind Enter key to load
        entry.bind('<Return>', lambda e: load_sheet())
        
    def copy_attendance_to_clipboard(self):
        """Copy attendance data to clipboard in the specified format"""
        try:
            # Get matched participants with roll numbers from both current session and persistent records
            matched_participants = []
            
            # Add current session records
            for name, match_data in self.matcher.matched_records.items():
                if match_data.get('status') == 'matched' and match_data.get('roll') and match_data.get('roll') != 'N/A':
                    try:
                        roll_num = int(match_data['roll'])
                        matched_participants.append(roll_num)
                    except ValueError:
                        # Skip if roll number is not a valid integer
                        continue
            
            # Add persistent records
            for name, match_data in self.matcher.persistent_records.items():
                if match_data.get('status') == 'matched' and match_data.get('roll') and match_data.get('roll') != 'N/A':
                    try:
                        roll_num = int(match_data['roll'])
                        if roll_num not in matched_participants:  # Avoid duplicates
                            matched_participants.append(roll_num)
                    except ValueError:
                        # Skip if roll number is not a valid integer
                        continue
            
            # Sort the roll numbers
            matched_participants.sort()
            
            # Format the data as specified
            date_str = datetime.now().strftime("%d.%m.%y")
            code = self.course_code_var.get()
            roll_str = f"({','.join(map(str, matched_participants))})"
            
            # Create the formatted text
            formatted_text = f"Date: {date_str}\n\nCode: {code}\n\nROLL:\n\n{roll_str}"
            
            # Copy to clipboard
            pyperclip.copy(formatted_text)
            self.log("Attendance data copied to clipboard", "success")
            
        except Exception as e:
            self.log(f"Error copying to clipboard: {e}")
            messagebox.showerror("Error", f"Failed to copy to clipboard:\n{e}")
                
    # Removed auto_detect_region and manual_select_region methods as they are no longer needed
    # with the new Zoom meeting approach
        
    def start_tracking(self):
        """Start tracking by joining Zoom meeting"""
        if not ZOOM_MEETING_AVAILABLE:
            messagebox.showerror("Error", "Zoom meeting functionality not available. Please install selenium and faker.")
            return
        
        # Get meeting details
        meeting_id = self.meeting_id_var.get().strip()
        passcode = self.passcode_var.get().strip()
        participants = self.participants_var.get()
        
        if not meeting_id:
            messagebox.showerror("Error", "Please enter a meeting ID.")
            return
        
        if not passcode:
            messagebox.showerror("Error", "Please enter a passcode.")
            return
        
        if participants < 1:
            messagebox.showerror("Error", "Please enter a valid number of participants.")
            return
        
        try:
            self.stop_event.clear()
            self.meeting_active = True
            self.is_tracking = True
            self.btn_start.config(state=tk.DISABLED)
            self.btn_stop.config(state=tk.NORMAL)
            # Reset refresh counter when starting new session
            self.refresh_count = 0
            self.update_refresh_counter()
            self.log(f"Joining Zoom meeting {meeting_id} with {participants} participants...", "info")
            self.status_bar.config(text="Joining Zoom meeting...")
            
            # Start meeting thread
            self.meeting_thread = threading.Thread(
                target=self._join_zoom_meeting,
                args=(meeting_id, passcode, participants),
                daemon=True
            )
            self.meeting_thread.start()
            
            # Start participant monitoring thread
            monitor_thread = threading.Thread(target=self._monitor_participants, daemon=True)
            monitor_thread.start()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to start tracking:\n{e}")
            
    def stop_tracking(self):
        """Stop tracking"""
        self.stop_event.set()
        self.meeting_active = False
        self.is_tracking = False
        self.btn_start.config(state=tk.NORMAL)
        self.btn_stop.config(state=tk.DISABLED)
        self.log("Meeting stopped", "warning")
        self.status_bar.config(text="Ready")
        self.generate_report()
        
    def reset_data(self):
        """Reset all data"""
        if messagebox.askyesno("Confirm", "Reset all attendance data?"):
            self.matcher.matched_records = {}
            self.tree.delete(*self.tree.get_children())
            self.update_stats(0, 0, 0)
            # Reset refresh counter
            self.refresh_count = 0
            self.update_refresh_counter()
            self.log("Data reset", "warning")
            
    # Removed update_tile_height method as it's no longer needed
    # with the new Zoom meeting approach
        
    def update_threshold(self, value):
        """Update match threshold"""
        val = int(float(value))
        self.threshold_label.config(text=str(val))
        self.matcher.threshold = val
        
    def toggle_continuous_save(self):
        """Toggle continuous save feature"""
        self.continuous_save_enabled = self.continuous_save_var.get()
        if self.continuous_save_enabled:
            self.log("Continuous save enabled", "success")
        else:
            self.log("Continuous save disabled", "warning")
        
    # Removed on_tracker_update method as it's no longer needed
    # with the new Zoom meeting approach
            
    def update_participant_list(self, matches):
        """Update participant treeview - show all matched participants"""
        self.tree.delete(*self.tree.get_children())
        
        matched_count = 0
        total_count = 0
        
        # Show all participants, but highlight matched ones
        for name, match in matches.items():
            status = match['status']
            total_count += 1
            
            # Display all participants, but differentiate matched vs unmatched
            if status == 'matched':
                matched_count += 1
                tag = 'matched'
                
                # Insert matched participants with roll numbers
                self.tree.insert('', tk.END, values=(
                    name,
                    match['roll'] if match['roll'] and match['roll'] != 'N/A' else 'N/A',
                    f"{match['confidence']:.0f}%",
                    'Matched'
                ), tags=(tag,))
            else:
                # Show unmatched participants with basic info
                tag = 'unmatched'
                self.tree.insert('', tk.END, values=(
                    name,
                    'N/A',
                    '0%',
                    'Unmatched'
                ), tags=(tag,))
        
        # Tags are configured in create_live_tab method
        
        # Update status bar with participant count
        self.status_bar.config(text=f"Ready - {total_count} participants detected, {matched_count} matched")
        
        # Update stats to show all participants
        self.update_stats(matched_count, total_count, matched_count)
        
    def update_stats(self, active, total, matched):
        """Update statistics"""
        self.stat_active.config(text=f"Active: {active}")
        self.stat_total.config(text=f"Total Detected: {total}")
        self.stat_matched.config(text=f"Matched: {matched}")
        
    def update_refresh_counter(self):
        """Update the refresh counter display"""
        self.stat_refresh.config(text=f"Refreshes: {self.refresh_count}")
        
    def log(self, message, color=None):
        """Add to event log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.config(state=tk.NORMAL)
        
        # Configure tags for different log levels
        self.log_text.tag_config("info", foreground="#2c3e50")
        self.log_text.tag_config("success", foreground="#27ae60")
        self.log_text.tag_config("warning", foreground="#f39c12")
        self.log_text.tag_config("error", foreground="#e74c3c")
        
        if color:
            # Use predefined color tags
            tag = color
            self.log_text.insert(tk.END, f"[{timestamp}] {message}\n", tag)
        else:
            # Default info color
            self.log_text.insert(tk.END, f"[{timestamp}] {message}\n", "info")
        
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        
    def save_continuous_report(self):
        """Save report continuously to a single file"""
        try:
            # Use a single filename
            filename = "attendance_report.txt"
            
            # Generate report content using combined records
            stats = self.matcher.get_statistics() if self.roll_file_loaded else None
            
            report = "=" * 60 + "\n"
            report += "ZOOM ATTENDANCE REPORT\n"
            report += "=" * 60 + "\n\n"
            report += f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            report += f"Made by: Fahad Akash\n\n"
            
            if stats:
                report += f"Total Detected: {stats['total_detected']}\n"
                report += f"Matched: {stats['matched']}\n"
                report += f"Unknown: {stats['unknown']}\n"
                report += f"Match Rate: {stats['match_rate']:.1f}%\n\n"
            
            report += "-" * 60 + "\n"
            report += "ATTENDANCE DETAILS\n"
            report += "-" * 60 + "\n\n"
            
            # Get matched participants from both current session and persistent records
            matched_participants = []
            
            # Add current session records
            for name, match_data in self.matcher.matched_records.items():
                if match_data.get('status') == 'matched' and match_data.get('roll') and match_data.get('roll') != 'N/A':
                    matched_participants.append((name, match_data['roll']))
            
            # Add persistent records
            for name, match_data in self.matcher.persistent_records.items():
                if match_data.get('status') == 'matched' and match_data.get('roll') and match_data.get('roll') != 'N/A':
                    # Check if this participant is already in the list to avoid duplicates
                    if not any(roll == match_data['roll'] for _, roll in matched_participants):
                        matched_participants.append((name, match_data['roll']))
            
            # Sort by roll number
            matched_participants.sort(key=lambda x: x[1])
            
            for name, roll in matched_participants:
                report += f"  • {name:<30} Roll: {roll}\n"
            
            # Save to file (overwrite each time)
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(report)
            
            self.log(f"Report updated continuously: {filename}", "success")
            
        except Exception as e:
            self.log(f"Error saving continuous report: {e}", "error")
        
    def generate_report(self):
        """Generate session report with only matched participants"""
        stats = self.matcher.get_statistics() if self.roll_file_loaded else None
        
        report = "=" * 60 + "\n"
        report += "ZOOM ATTENDANCE REPORT\n"
        report += "=" * 60 + "\n\n"
        report += f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
        
        if stats:
            report += f"Total Detected: {stats['total_detected']}\n"
            report += f"Matched: {stats['matched']}\n"
            report += f"Unknown: {stats['unknown']}\n"
            report += f"Match Rate: {stats['match_rate']:.1f}%\n\n"
        
        report += "-" * 60 + "\n"
        report += "ATTENDANCE DETAILS\n"
        report += "-" * 60 + "\n\n"
        
        # Get matched participants from both current session and persistent records
        matched_participants = []
        
        # Add current session records
        for name, match_data in self.matcher.matched_records.items():
            if match_data.get('status') == 'matched' and match_data.get('roll') and match_data.get('roll') != 'N/A':
                matched_participants.append((name, match_data['roll']))
        
        # Add persistent records
        for name, match_data in self.matcher.persistent_records.items():
            if match_data.get('status') == 'matched' and match_data.get('roll') and match_data.get('roll') != 'N/A':
                # Check if this participant is already in the list to avoid duplicates
                if not any(roll == match_data['roll'] for _, roll in matched_participants):
                    matched_participants.append((name, match_data['roll']))
        
        # Sort by roll number
        matched_participants.sort(key=lambda x: x[1])
        
        for name, roll in matched_participants:
            report += f"  • {name:<30} Roll: {roll}\n"
        
        self.report_text.delete(1.0, tk.END)
        self.report_text.insert(1.0, report)
        
    def export_excel(self):
        """Export only matched participants to Excel"""
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        
        if filepath:
            try:
                # Get data from both current session and persistent records
                all_data = self.matcher.export_attendance()
                matched_data = [row for row in all_data if row['Status'] == 'Matched' and row['Roll Number'] != 'N/A']
                
                df = pd.DataFrame(matched_data)
                df.to_excel(filepath, index=False)
                messagebox.showinfo("Success", f"Exported {len(matched_data)} matched participants to:\n{filepath}")
                self.log(f"Exported {len(matched_data)} matched participants to Excel: {filepath}", "success")
            except Exception as e:
                messagebox.showerror("Error", f"Export failed:\n{e}")
                
    def export_csv(self):
        """Export only matched participants to CSV"""
        filepath = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")]
        )
        
        if filepath:
            try:
                # Get data from both current session and persistent records
                all_data = self.matcher.export_attendance()
                matched_data = [row for row in all_data if row['Status'] == 'Matched' and row['Roll Number'] != 'N/A']
                
                df = pd.DataFrame(matched_data)
                df.to_csv(filepath, index=False)
                messagebox.showinfo("Success", f"Exported {len(matched_data)} matched participants to:\n{filepath}")
                self.log(f"Exported {len(matched_data)} matched participants to CSV: {filepath}", "success")
            except Exception as e:
                messagebox.showerror("Error", f"Export failed:\n{e}")
                
    def refresh_zoom_window(self):
        """Refresh functionality - not applicable with new approach"""
        messagebox.showinfo("Info", "Refresh not needed with the new Zoom meeting approach.\n\nJust enter your meeting details and click 'Join Meeting'.")
    
    def _join_zoom_meeting(self, meeting_id, passcode, num_participants):
        """Join Zoom meeting with specified parameters"""
        try:
            drivers = []
            
            self.log(f"Starting {num_participants} participants for meeting {meeting_id}")
            
            for i in range(num_participants):
                if self.stop_event.is_set():
                    break
                    
                try:
                    # Create WebDriver
                    options = webdriver.ChromeOptions()
                    options.add_argument(f'user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.83 Safari/537.36')
                    options.add_experimental_option("detach", True)
                    options.add_argument("--window-size=1920,1080")
                    options.add_argument('--no-sandbox')
                    options.add_argument('--disable-dev-shm-usage')
                    options.add_argument('--ignore-certificate-errors')
                    options.add_argument('--allow-running-insecure-content')
                    options.add_argument('allow-file-access-from-files')
                    options.add_argument('use-fake-device-for-media-stream')
                    options.add_argument('use-fake-ui-for-media-stream')
                    options.add_argument("--disable-extensions")
                    options.add_argument("--proxy-server='direct://'")
                    options.add_argument("--proxy-bypass-list=*")
                    options.add_argument("--start-maximized")
                    
                    # Disable audio and video for silent participation
                    options.add_argument("--use-fake-ui-for-media-stream")
                    options.add_argument("--disable-audio-output")
                    options.add_argument("--disable-background-media-suspend")
                    options.add_argument("--disable-renderer-backgrounding")
                    options.add_argument("--autoplay-policy=no-user-gesture-required")
                    
                    # Additional options for Zoom meetings
                    options.add_argument("--disable-features=TranslateUI")
                    options.add_argument("--disable-ipc-flooding-protection")
                    options.add_argument("--disable-backgrounding-occluded-windows")
                    options.add_argument("--disable-breakpad")
                    
                    # Try to use webdriver-manager
                    try:
                        service = Service(ChromeDriverManager().install())
                        driver = webdriver.Chrome(service=service, options=options)
                    except Exception as e:
                        self.log(f"Error with ChromeDriver manager: {e}. Trying without Service...")
                        try:
                            driver = webdriver.Chrome(options=options)
                        except Exception as e2:
                            self.log(f"FATAL: Could not start Chrome: {e2}")
                            return
                    
                    drivers.append(driver)
                    
                    # Navigate to meeting
                    driver.get(f'https://app.zoom.us/wc/join/{meeting_id}')
                    time.sleep(2)
                    
                    # Fill passcode if present
                    try:
                        inp2 = WebDriverWait(driver, 3).until(ec.presence_of_element_located((By.ID, 'input-for-pwd')))
                        inp2.clear()
                        inp2.send_keys(passcode)
                    except Exception:
                        pass
                    
                    # Fill display name
                    user_name = f"Batch-72-{i+1:02d}"
                    try:
                        inp = WebDriverWait(driver, 5).until(ec.presence_of_element_located((By.ID, 'input-for-name')))
                        inp.clear()
                        inp.send_keys(f"{user_name}")
                    except Exception:
                        # fallback to common text input
                        try:
                            inp = driver.find_element(By.CSS_SELECTOR, 'input[type="text"]')
                            inp.clear()
                            inp.send_keys(f"{user_name}")
                        except Exception:
                            pass
                    
                    # Automatically mute audio and video before joining
                    try:
                        # Mute microphone using JavaScript
                        driver.execute_script("""
                            // Try to find and click mute microphone button
                            var micButtons = document.querySelectorAll('[class*="audio"], [aria-label*="microphone"], [aria-label*="Microphone"]');
                            for (var i = 0; i < micButtons.length; i++) {
                                var button = micButtons[i];
                                if (button.textContent.includes('Mute') || button.textContent.includes('mute') || 
                                    button.getAttribute('aria-label')?.includes('Mute') || button.getAttribute('aria-label')?.includes('mute')) {
                                    if (!button.disabled) {
                                        button.click();
                                        return true;
                                    }
                                }
                            }
                            return false;
                        """)
                        self.log(f"Attempted to mute microphone for {user_name}")
                    except Exception as e:
                        self.log(f"Could not mute microphone for {user_name}: {e}")
                    
                    try:
                        # Turn off video using JavaScript
                        driver.execute_script("""
                            // Try to find and click turn off video button
                            var videoButtons = document.querySelectorAll('[class*="video"], [aria-label*="video"], [aria-label*="Video"]');
                            for (var i = 0; i < videoButtons.length; i++) {
                                var button = videoButtons[i];
                                if (button.textContent.includes('Turn off') || button.textContent.includes('turn off') || 
                                    button.getAttribute('aria-label')?.includes('Turn off') || button.getAttribute('aria-label')?.includes('turn off')) {
                                    if (!button.disabled) {
                                        button.click();
                                        return true;
                                    }
                                }
                            }
                            return false;
                        """)
                        self.log(f"Attempted to turn off video for {user_name}")
                    except Exception as e:
                        self.log(f"Could not turn off video for {user_name}: {e}")
                    
                    try:
                        btn2 = WebDriverWait(driver, 5).until(ec.element_to_be_clickable((By.CLASS_NAME, 'zm-btn')))
                        try:
                            btn2.click()
                        except Exception:
                            driver.execute_script("arguments[0].click();", btn2)
                        time.sleep(1)
                    except Exception as e:
                        pass
                    
                    try:
                        WebDriverWait(driver, 10).until(ec.element_to_be_clickable((By.CSS_SELECTOR, '#preview-audio-control-button')))
                        audio_btn = driver.find_element(By.CSS_SELECTOR, '#preview-audio-control-button')
                        try:
                            audio_btn.click()
                        except Exception:
                            driver.execute_script("arguments[0].click();", audio_btn)
                        time.sleep(0.5)
                    except:
                        pass
                    
                    try:
                        btn3 = WebDriverWait(driver, 10).until(ec.element_to_be_clickable((By.CLASS_NAME, "preview-join-button")))
                        try:
                            btn3.click()
                        except Exception:
                            driver.execute_script("arguments[0].click();", btn3)
                        time.sleep(1)
                    except Exception as e:
                        pass
                    
                    try:
                        driver.find_element(By.XPATH, '//*[@id="voip-tab"]/div/button').click()
                    except Exception as e:
                        pass
                    time.sleep(0.5)
                    
                    # Mute audio and video after joining
                    try:
                        # Mute microphone after joining
                        mute_buttons = driver.find_elements(By.CSS_SELECTOR, '[aria-label*="mute"], [aria-label*="Mute"], [class*="mic"], [class*="Mic"]')
                        for button in mute_buttons:
                            if 'mute' in button.get_attribute('aria-label').lower() or 'mic' in button.get_attribute('class').lower():
                                if button.is_enabled() and button.is_displayed():
                                    button.click()
                                    self.log(f"Muted microphone for {user_name} after joining")
                                    break
                    except Exception as e:
                        self.log(f"Could not mute microphone after joining for {user_name}: {e}")
                    
                    try:
                        # Turn off video after joining
                        video_buttons = driver.find_elements(By.CSS_SELECTOR, '[aria-label*="video"], [aria-label*="Video"], [class*="video"], [class*="Video"]')
                        for button in video_buttons:
                            if 'video' in button.get_attribute('aria-label').lower() and 'off' in button.get_attribute('aria-label').lower():
                                if button.is_enabled() and button.is_displayed():
                                    button.click()
                                    self.log(f"Turned off video for {user_name} after joining")
                                    break
                            elif 'video' in button.get_attribute('class').lower() and 'off' in button.get_attribute('class').lower():
                                if button.is_enabled() and button.is_displayed():
                                    button.click()
                                    self.log(f"Turned off video for {user_name} after joining")
                                    break
                    except Exception as e:
                        self.log(f"Could not turn off video after joining for {user_name}: {e}")
                    
                    # Use JavaScript to ensure audio and video are muted
                    try:
                        driver.execute_script("""
                            // Mute audio
                            if (typeof APP !== 'undefined' && APP.conference) {
                                APP.conference.toggleAudioMuted();
                            }
                            
                            // Turn off video
                            if (typeof APP !== 'undefined' && APP.conference) {
                                APP.conference.toggleVideoMuted();
                            }
                            
                            // Alternative Zoom-specific muting
                            var muteButton = document.querySelector('[aria-label*="Mute" i]');
                            if (muteButton) muteButton.click();
                            
                            var videoButton = document.querySelector('[aria-label*="Video" i][aria-label*="Off" i]');
                            if (videoButton) videoButton.click();
                        """)
                        self.log(f"Ensured silent participation for {user_name}")
                    except Exception as e:
                        self.log(f"Could not ensure silent participation for {user_name}: {e}")
                    
                    self.log(f"Participant {i+1} ({user_name}) joined meeting")
                    
                except Exception as e:
                    self.log(f"Error joining participant {i+1}: {e}")
                    continue
            
            # Keep drivers alive and periodically check for participants
            check_interval = 0
            while not self.stop_event.is_set():
                # Fetch participants from one of the drivers
                if drivers and check_interval <= 0:
                    try:
                        participants = self._fetch_participants(drivers[0])
                        if participants:
                            # Put participants in queue for processing
                            self.participants_queue.put(participants)
                            self.log(f"Fetched {len(participants)} participants")
                            
                            # Increment refresh counter and update display
                            self.refresh_count += 1
                            self.root.after(0, self.update_refresh_counter)
                    except Exception as e:
                        self.log(f"Error fetching participants: {e}")
                    
                    # Reset check interval (check every 10 seconds)
                    check_interval = 10
                
                # Decrement check interval
                check_interval -= 1
                time.sleep(1)
            
            # Quit all drivers when stopping
            for driver in drivers:
                try:
                    driver.quit()
                except Exception:
                    pass
                    
            self.log("Meeting ended")
            
        except Exception as e:
            self.log(f"Error in meeting thread: {e}")
    
    def _fetch_participants(self, driver):
        """Fetch participants from Zoom meeting"""
        try:
            time.sleep(3)  # Wait for meeting to fully load
            
            # Try to open Participants panel
            open_selectors = [
                'button[aria-label*="Participants"]',
                'button[aria-label*="participants"]',
                'button[title*="Participants"]',
                'button[title*="participants"]',
                'button[aria-label*="Participants and chat"]',
                'button[aria-label="Participants"]',
                'button.participants-tab',
            ]
            
            for sel in open_selectors:
                try:
                    btn = driver.find_element(By.CSS_SELECTOR, sel)
                    try:
                        btn.click()
                    except:
                        driver.execute_script("arguments[0].click();", btn)
                    time.sleep(1)
                    break
                except Exception:
                    pass
            
            # Try XPath selectors as fallback
            xpath_selectors = [
                '//button[contains(@aria-label, "Participant")]',
                '//*[contains(text(), "Participant")]',
            ]
            
            for xpath in xpath_selectors:
                try:
                    btn = driver.find_element(By.XPATH, xpath)
                    try:
                        btn.click()
                    except:
                        driver.execute_script("arguments[0].click();", btn)
                    time.sleep(1)
                    break
                except Exception:
                    pass
            
            # Candidate selectors for participant name elements
            name_selectors = [
                '#zmu-portal-dropdown-participant-list li',  # Zoom web client structure
                '.participants-section-container_wrapper li',
                'div[class*="participant"] span',
                'div[class*="participant-name"]',
                '.participants-item__display-name',
                '.participants-item__name',
                '.participant-name',
                '.name',
                '.participants-list li',
                '.participants-list div',
                '[data-testid*="participant"]',
                '.zm-participant-name',
            ]
            
            names = []
            for sel in name_selectors:
                try:
                    elems = driver.find_elements(By.CSS_SELECTOR, sel)
                    for e in elems:
                        text = e.text.strip()
                        if text and text not in names and len(text) > 1:
                            # Clean up the participant name
                            # Remove (Host, me) and similar annotations
                            import re
                            text = re.sub(r'\(.*?(?:Host|Me|me).*?\)', '', text, flags=re.IGNORECASE).strip()
                            
                            # Filter out unwanted UI elements
                            unwanted_patterns = [
                                'Unmute', 'start Video', 'Participants', 'chat', 'Reactions', 
                                'Share Screen', 'more', 'leave', 'Pleader', 'upgrade your browser',
                                'update your browder', 'Speaker', 'Gallery View', 'Participant (',
                                'Mute', 'Turn off', 'NEW', 'Invite', 'Record', 'Security', 'Manage Participants',
                                'Stop Video', 'Batch-72-01'
                            ]
                            
                            is_unwanted = False
                            for pattern in unwanted_patterns:
                                if pattern.lower() in text.lower():
                                    is_unwanted = True
                                    break
                            
                            # Also filter out text with numbers in parentheses like "(3)"
                            if re.search(r'\(\s*\d+\s*\)', text):
                                is_unwanted = True
                                
                            if not is_unwanted and text and text not in names:
                                names.append(text)
                    if names:
                        break
                except Exception:
                    pass
            
            return names
            
        except Exception as e:
            self.log(f"Error fetching participants: {e}")
            return []
    
    def _monitor_participants(self):
        """Monitor participant queue and update UI continuously"""
        while not self.stop_event.is_set():
            try:
                # Process all available participant data in the queue
                processed = False
                while not self.participants_queue.empty():
                    try:
                        participants = self.participants_queue.get_nowait()
                        if participants:
                            # Match names if roll file loaded
                            if self.roll_file_loaded:
                                matches = self.matcher.match_batch(participants)
                            else:
                                matches = {name: {'matched_name': name, 'roll': 'N/A', 'confidence': 0, 'status': 'no_db'} 
                                          for name in participants}
                            
                            # Update UI
                            self.root.after(0, lambda m=matches: self.update_participant_list(m))
                            
                            # Log event
                            self.root.after(0, lambda p=len(participants): self.log(f"Participants updated: {p} detected"))
                            
                            # Continuous save if enabled
                            if self.continuous_save_enabled:
                                self.root.after(0, self.save_continuous_report)
                            
                            processed = True
                    except queue.Empty:
                        break
                
                # If we processed data, add a small delay to prevent excessive CPU usage
                if processed:
                    time.sleep(0.1)
                else:
                    # If no data, wait a bit longer
                    time.sleep(0.5)
                    
            except Exception as e:
                self.log(f"Error in participant monitor: {e}")
                time.sleep(1)  # Wait longer on error to prevent spam
    
    def show_help(self):
        """Show user guide dialog"""
        help_text = (
            "Zoom Attendance System - User Guide\n\n"
            "1. SETUP TAB:\n"
            "   - Load student data from a text file or Google Sheet\n"
            "   - Enter Zoom meeting details (ID, passcode, participants)\n"
            "   - Adjust match threshold and enable continuous save if needed\n\n"
            "2. LIVE TRACKING TAB:\n"
            "   - Click '▶ Join Meeting' to start tracking\n"
            "   - View real-time participant updates\n"
            "   - Click '⏹ Stop' to end tracking\n\n"
            "3. REPORTS TAB:\n"
            "   - View attendance report\n"
            "   - Export to Excel/CSV or copy to clipboard\n\n"
            "Keyboard Shortcuts:\n"
            "   Ctrl+O - Load Roll Numbers\n"
            "   Ctrl+E - Export to Excel\n"
            "   Ctrl+Shift+E - Export to CSV\n"
            "   Ctrl+Q - Exit Application"
        )
        messagebox.showinfo("User Guide", help_text)
        
    def show_about(self):
        """Show about dialog"""
        messagebox.showinfo("About", 
                           "Zoom Attendance System v1.0\n\n"
                           "Automatically track Zoom meeting attendance\n"
                           "with roll number matching.\n\n"
                           "Made by: Fahad Akash\n\n"
                           "Built with Python, Tkinter, Selenium, and RapidFuzz")


def main():
    root = ttk.Window(themename="superhero")
    app = AttendanceApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
