#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ví dụ đơn giản: Tích hợp Auto-Update vào Desktop App

Chạy: python example_simple_desktop_integration.py

Tính năng:
- Desktop app với Tkinter
- Menu "Check for Updates"
- Tự động kiểm tra update khi khởi động
- Hiển thị changelog
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
from pathlib import Path
import sys

# Import updater
sys.path.insert(0, str(Path(__file__).parent))
from updater import AutoUpdater, VersionManager


class SimpleApp:
    """Ví dụ ứng dụng desktop đơn giản"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("HocBaPdfToExcel - Example")
        self.root.geometry("600x400")
        
        # Tạo UI
        self.create_ui()
        
        # Kiểm tra update khi khởi động (background)
        self.check_update_on_startup()
    
    def create_ui(self):
        """Tạo giao diện"""
        
        # Menu bar
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="Check for Updates", command=self.on_check_update)
        help_menu.add_separator()
        help_menu.add_command(label="About", command=self.show_about)
        
        # Main content
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(
            main_frame,
            text="HocBaPdf To Excel Converter",
            font=("Arial", 16, "bold")
        )
        title_label.grid(row=0, column=0, columnspan=2, pady=10)
        
        # Version info
        version = VersionManager.get_current_version()
        version_label = ttk.Label(
            main_frame,
            text=f"Current Version: {version}",
            font=("Arial", 10)
        )
        version_label.grid(row=1, column=0, columnspan=2, pady=5)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=20)
        
        ttk.Button(
            button_frame,
            text="Select Input Folder",
            command=self.select_input_folder
        ).pack(fill=tk.X, pady=5)
        
        ttk.Button(
            button_frame,
            text="Select Output Folder",
            command=self.select_output_folder
        ).pack(fill=tk.X, pady=5)
        
        ttk.Button(
            button_frame,
            text="Start Conversion",
            command=self.start_conversion
        ).pack(fill=tk.X, pady=5)
        
        # Status text
        status_frame = ttk.LabelFrame(main_frame, text="Status", padding="5")
        status_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        self.status_text = tk.Text(status_frame, height=10, width=70)
        self.status_text.pack(fill=tk.BOTH, expand=True)
        
        self.log_message("Ready. Select input folder to start.")
    
    def log_message(self, msg):
        """Log message to status text"""
        self.status_text.insert(tk.END, msg + "\n")
        self.status_text.see(tk.END)
        self.root.update_idletasks()
    
    def select_input_folder(self):
        """Select input folder"""
        folder = filedialog.askdirectory(title="Select Input Folder")
        if folder:
            self.log_message(f"Input: {folder}")
    
    def select_output_folder(self):
        """Select output folder"""
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.log_message(f"Output: {folder}")
    
    def start_conversion(self):
        """Start conversion (dummy)"""
        self.log_message("Conversion started...")
        self.log_message("Processing...")
        self.log_message("Done!")
    
    def on_check_update(self):
        """Handle check update button"""
        # Run in thread to not block UI
        thread = threading.Thread(target=self._run_update_check, daemon=True)
        thread.start()
    
    def _run_update_check(self):
        """Check for updates"""
        try:
            self.log_message("\n🔍 Checking for updates...")
            updater = AutoUpdater()
            
            has_update, release = updater.checker.has_update()
            
            if has_update:
                new_version = release['tag_name'].lstrip('v')
                self.log_message(f"✅ New version available: {new_version}")
                
                # Ask user
                if messagebox.askyesno(
                    "Update Available",
                    f"New version {new_version} available.\n\nUpdate now?"
                ):
                    self.log_message("📥 Downloading update...")
                    updater.check_and_update(show_dialog=False, callback=self._update_callback)
            else:
                self.log_message("✅ You have the latest version")
                messagebox.showinfo("No Update", "You have the latest version")
        
        except Exception as e:
            self.log_message(f"❌ Error: {e}")
            messagebox.showerror("Error", f"Error checking for updates:\n{e}")
    
    def _update_callback(self, status, message):
        """Callback from updater"""
        if status == 'no_update':
            self.log_message(f"✅ {message}")
        elif status == 'update_available':
            self.log_message(f"📦 {message}")
        elif status == 'download_complete':
            self.log_message(f"✅ {message}")
        elif status == 'update_complete':
            self.log_message(f"✨ {message}")
            messagebox.showinfo(
                "Update Complete",
                "Update completed successfully!\n\nPlease restart the application."
            )
        elif status == 'error':
            self.log_message(f"❌ {message}")
    
    def check_update_on_startup(self):
        """Check for updates on startup (background)"""
        def run():
            try:
                updater = AutoUpdater()
                has_update, _ = updater.checker.has_update()
                
                if has_update:
                    # Show notification
                    self.root.after(0, lambda: messagebox.showinfo(
                        "Update Available",
                        "A new version is available!\n\n"
                        "Go to Help → Check for Updates"
                    ))
            except:
                pass  # Silently fail on startup
        
        thread = threading.Thread(target=run, daemon=True)
        thread.start()
    
    def show_about(self):
        """Show about dialog"""
        version = VersionManager.get_current_version()
        messagebox.showinfo(
            "About HocBaPdfToExcel",
            f"HocBaPdfToExcel v{version}\n\n"
            "PDF to Excel Converter for Vietnamese Student Records\n\n"
            "Author: Võ Thành Khải\n\n"
            "Help → Check for Updates to get the latest version"
        )


if __name__ == '__main__':
    root = tk.Tk()
    app = SimpleApp(root)
    root.mainloop()
