"""
GUI Module - Zoom Attendance System
Main application interface using Tkinter
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
from datetime import datetime
import pandas as pd
from tracker import ZoomTracker
from matcher import RollMatcher


class AttendanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Zoom Attendance System")
        self.root.geometry("900x700")
        
        # Initialize components
        self.tracker = ZoomTracker(callback=self.on_tracker_update)
        self.matcher = RollMatcher(threshold=75)
        
        # State
        self.is_tracking = False
        self.roll_file_loaded = False
        
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
        file_menu.add_command(label="Load Roll Numbers", command=self.load_roll_file)
        file_menu.add_separator()
        file_menu.add_command(label="Export to Excel", command=self.export_excel)
        file_menu.add_command(label="Export to CSV", command=self.export_csv)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)
        
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
        # Roll number file section
        frame_roll = ttk.LabelFrame(self.setup_tab, text="Roll Number Database", padding=10)
        frame_roll.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(frame_roll, text="Load file with format: 'Name Roll'").pack(anchor=tk.W)
        ttk.Label(frame_roll, text="Example: Fahad Akash 08", font=('Arial', 9, 'italic')).pack(anchor=tk.W)
        
        btn_frame = ttk.Frame(frame_roll)
        btn_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(btn_frame, text="Browse File", command=self.load_roll_file).pack(side=tk.LEFT, padx=5)
        self.roll_status = ttk.Label(btn_frame, text="No file loaded", foreground="red")
        self.roll_status.pack(side=tk.LEFT, padx=5)
        
        # Region selection
        frame_region = ttk.LabelFrame(self.setup_tab, text="Capture Region", padding=10)
        frame_region.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(frame_region, text="Auto-Detect Zoom Window", 
                  command=self.auto_detect_region).pack(fill=tk.X, pady=5)
        ttk.Button(frame_region, text="Manual Selection (Drag)", 
                  command=self.manual_select_region).pack(fill=tk.X, pady=5)
        
        self.region_status = ttk.Label(frame_region, text="Region not set", foreground="orange")
        self.region_status.pack(pady=5)
        
        # Settings
        frame_settings = ttk.LabelFrame(self.setup_tab, text="Settings", padding=10)
        frame_settings.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(frame_settings, text="Tile Height (pixels):").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.tile_height_var = tk.IntVar(value=70)
        ttk.Scale(frame_settings, from_=30, to=150, variable=self.tile_height_var, 
                 orient=tk.HORIZONTAL, command=self.update_tile_height).grid(row=0, column=1, sticky=tk.EW, padx=5)
        self.tile_label = ttk.Label(frame_settings, text="70")
        self.tile_label.grid(row=0, column=2)
        
        ttk.Label(frame_settings, text="Match Threshold (%):").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.threshold_var = tk.IntVar(value=75)
        ttk.Scale(frame_settings, from_=50, to=100, variable=self.threshold_var, 
                 orient=tk.HORIZONTAL, command=self.update_threshold).grid(row=1, column=1, sticky=tk.EW, padx=5)
        self.threshold_label = ttk.Label(frame_settings, text="75")
        self.threshold_label.grid(row=1, column=2)
        
        frame_settings.columnconfigure(1, weight=1)
        
    def create_live_tab(self):
        """Live tracking tab"""
        # Controls
        control_frame = ttk.Frame(self.live_tab)
        control_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.btn_start = ttk.Button(control_frame, text="▶ Start Tracking", 
                                     command=self.start_tracking, style="Accent.TButton")
        self.btn_start.pack(side=tk.LEFT, padx=5)
        
        self.btn_stop = ttk.Button(control_frame, text="⏹ Stop", 
                                    command=self.stop_tracking, state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, padx=5)
        
        self.btn_reset = ttk.Button(control_frame, text="Reset Data", command=self.reset_data)
        self.btn_reset.pack(side=tk.LEFT, padx=5)
        
        # Stats
        stats_frame = ttk.LabelFrame(self.live_tab, text="Statistics", padding=10)
        stats_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.stat_active = ttk.Label(stats_frame, text="Active: 0", font=('Arial', 12, 'bold'))
        self.stat_active.pack(side=tk.LEFT, padx=20)
        
        self.stat_total = ttk.Label(stats_frame, text="Total Detected: 0", font=('Arial', 12))
        self.stat_total.pack(side=tk.LEFT, padx=20)
        
        self.stat_matched = ttk.Label(stats_frame, text="Matched: 0", font=('Arial', 12))
        self.stat_matched.pack(side=tk.LEFT, padx=20)
        
        # Participant list
        list_frame = ttk.LabelFrame(self.live_tab, text="Current Participants", padding=10)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Treeview
        columns = ('Name', 'Roll', 'Confidence', 'Status')
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=15)
        
        self.tree.heading('Name', text='Detected Name')
        self.tree.heading('Roll', text='Roll Number')
        self.tree.heading('Confidence', text='Match %')
        self.tree.heading('Status', text='Status')
        
        self.tree.column('Name', width=250)
        self.tree.column('Roll', width=100)
        self.tree.column('Confidence', width=100)
        self.tree.column('Status', width=100)
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Event log
        log_frame = ttk.LabelFrame(self.live_tab, text="Event Log", padding=5)
        log_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=5, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
    def create_report_tab(self):
        """Reports tab"""
        frame = ttk.Frame(self.report_tab, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="Session Report", font=('Arial', 14, 'bold')).pack(pady=10)
        
        self.report_text = scrolledtext.ScrolledText(frame, height=20, font=('Courier', 10))
        self.report_text.pack(fill=tk.BOTH, expand=True, pady=10)
        
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(btn_frame, text="Refresh Report", command=self.generate_report).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Export Excel", command=self.export_excel).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Export CSV", command=self.export_csv).pack(side=tk.LEFT, padx=5)
        
    def create_status_bar(self):
        """Status bar at bottom"""
        self.status_bar = ttk.Label(self.root, text="Ready", relief=tk.SUNKEN, anchor=tk.W)
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
                self.log(f"Loaded {count} roll numbers from file")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file:\n{e}")
                
    def auto_detect_region(self):
        """Auto-detect Zoom window"""
        self.status_bar.config(text="Searching for Zoom window...")
        region = self.tracker.find_zoom_window()
        
        if region:
            self.tracker.set_region(region)
            self.region_status.config(text="✓ Region set (auto)", foreground="green")
            self.log("Zoom window detected automatically")
        else:
            messagebox.showwarning("Not Found", "Zoom window not found. Try manual selection.")
        
        self.status_bar.config(text="Ready")
        
    def manual_select_region(self):
        """Manual region selection"""
        self.status_bar.config(text="Click and drag to select region...")
        messagebox.showinfo("Manual Selection", 
                           "Click and drag with your mouse over the participant list area.\n\n"
                           "Release to confirm selection.")
        
        def select_thread():
            region = self.tracker.get_manual_region()
            if region:
                self.tracker.set_region(region)
                self.root.after(0, lambda: self.region_status.config(text="✓ Region set (manual)", foreground="green"))
                self.root.after(0, lambda: self.log("Region selected manually"))
            self.root.after(0, lambda: self.status_bar.config(text="Ready"))
        
        threading.Thread(target=select_thread, daemon=True).start()
        
    def start_tracking(self):
        """Start tracking"""
        if not self.tracker.region:
            messagebox.showwarning("Setup Required", "Please set capture region first!")
            return
        
        try:
            self.tracker.start()
            self.is_tracking = True
            self.btn_start.config(state=tk.DISABLED)
            self.btn_stop.config(state=tk.NORMAL)
            self.log("Tracking started")
            self.status_bar.config(text="Tracking active...")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to start tracking:\n{e}")
            
    def stop_tracking(self):
        """Stop tracking"""
        self.tracker.stop()
        self.is_tracking = False
        self.btn_start.config(state=tk.NORMAL)
        self.btn_stop.config(state=tk.DISABLED)
        self.log("Tracking stopped")
        self.status_bar.config(text="Ready")
        self.generate_report()
        
    def reset_data(self):
        """Reset all data"""
        if messagebox.askyesno("Confirm", "Reset all attendance data?"):
            self.tracker.reset()
            self.matcher.matched_records = {}
            self.tree.delete(*self.tree.get_children())
            self.update_stats(0, 0, 0)
            self.log("Data reset")
            
    def update_tile_height(self, value):
        """Update tile height"""
        val = int(float(value))
        self.tile_label.config(text=str(val))
        self.tracker.set_tile_height(val)
        
    def update_threshold(self, value):
        """Update match threshold"""
        val = int(float(value))
        self.threshold_label.config(text=str(val))
        self.matcher.threshold = val
        
    def on_tracker_update(self, participants, event_type, changed):
        """Callback from tracker"""
        # Match names if roll file loaded
        if self.roll_file_loaded:
            matches = self.matcher.match_batch(participants)
        else:
            matches = {name: {'matched_name': name, 'roll': 'N/A', 'confidence': 0, 'status': 'no_db'} 
                      for name in participants}
        
        # Update UI
        self.root.after(0, lambda: self.update_participant_list(matches))
        
        # Log events
        if event_type == 'joined' and changed:
            self.root.after(0, lambda: self.log(f"✓ Joined: {', '.join(changed)}", "green"))
        elif event_type == 'left' and changed:
            self.root.after(0, lambda: self.log(f"✗ Left: {', '.join(changed)}", "red"))
            
    def update_participant_list(self, matches):
        """Update participant treeview"""
        self.tree.delete(*self.tree.get_children())
        
        matched_count = 0
        for name, match in matches.items():
            status = match['status']
            if status == 'matched':
                matched_count += 1
                tag = 'matched'
            else:
                tag = 'unknown'
            
            self.tree.insert('', tk.END, values=(
                name,
                match['roll'],
                f"{match['confidence']:.0f}%",
                status.title()
            ), tags=(tag,))
        
        self.tree.tag_configure('matched', foreground='green')
        self.tree.tag_configure('unknown', foreground='orange')
        
        self.update_stats(len(matches), len(matches), matched_count)
        
    def update_stats(self, active, total, matched):
        """Update statistics"""
        self.stat_active.config(text=f"Active: {active}")
        self.stat_total.config(text=f"Total Detected: {total}")
        self.stat_matched.config(text=f"Matched: {matched}")
        
    def log(self, message, color=None):
        """Add to event log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.config(state=tk.NORMAL)
        
        if color:
            # Simple text coloring via tags
            tag = f"color_{color}"
            self.log_text.tag_config(tag, foreground=color)
            self.log_text.insert(tk.END, f"[{timestamp}] {message}\n", tag)
        else:
            self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        
    def generate_report(self):
        """Generate session report"""
        data = self.tracker.get_attendance_data()
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
        
        for name in sorted(data['current']):
            match = self.matcher.matched_records.get(name, {})
            roll = match.get('roll', 'N/A')
            report += f"  • {name:<30} Roll: {roll}\n"
        
        self.report_text.delete(1.0, tk.END)
        self.report_text.insert(1.0, report)
        
    def export_excel(self):
        """Export to Excel"""
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        
        if filepath:
            try:
                data = self.matcher.export_attendance()
                df = pd.DataFrame(data)
                df.to_excel(filepath, index=False)
                messagebox.showinfo("Success", f"Exported to:\n{filepath}")
                self.log(f"Exported to Excel: {filepath}")
            except Exception as e:
                messagebox.showerror("Error", f"Export failed:\n{e}")
                
    def export_csv(self):
        """Export to CSV"""
        filepath = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")]
        )
        
        if filepath:
            try:
                data = self.matcher.export_attendance()
                df = pd.DataFrame(data)
                df.to_csv(filepath, index=False)
                messagebox.showinfo("Success", f"Exported to:\n{filepath}")
                self.log(f"Exported to CSV: {filepath}")
            except Exception as e:
                messagebox.showerror("Error", f"Export failed:\n{e}")
                
    def show_about(self):
        """Show about dialog"""
        messagebox.showinfo("About", 
                           "Zoom Attendance System v1.0\n\n"
                           "Automatically track Zoom meeting attendance\n"
                           "with roll number matching.\n\n"
                           "Built with Python, Tkinter, OpenCV, and Tesseract OCR")


def main():
    root = tk.Tk()
    app = AttendanceApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
