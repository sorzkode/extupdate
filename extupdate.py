#!/usr/bin/env python3
"""
███████ ██   ██ ████████ ██    ██ ██████  ██████   █████  ████████ ███████ 
██       ██ ██     ██    ██    ██ ██   ██ ██   ██ ██   ██    ██    ██      
█████     ███      ██    ██    ██ ██████  ██   ██ ███████    ██    █████   
██       ██ ██     ██    ██    ██ ██      ██   ██ ██   ██    ██    ██      
███████ ██   ██    ██     ██████  ██      ██████  ██   ██    ██    ███████ 

Bulk Update/Change Excel File Extensions.
-
Author:
sorzkode
https://github.com/sorzkode

MIT License
Copyright (c) 2025
"""
import os
import sys
import shutil
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.font import Font
import threading
from typing import List, Optional, Dict
from datetime import datetime
import json

class ExtUpdateApp:
    """Main application class for Excel extension updater"""
    
    # Excel file extensions
    EXCEL_EXTENSIONS = [
        ".xls", ".xlsx", ".xlsm", ".xlsb", 
        ".xltx", ".xltm", ".xlt", ".xml"
    ]
    
    # Extension compatibility warnings
    COMPATIBILITY_WARNINGS = {
        ".xls": "Legacy format with limited features (65,536 rows, no XML)",
        ".xlsx": "Modern format, most compatible",
        ".xlsm": "Supports macros/VBA code",
        ".xlsb": "Binary format, faster but less compatible",
        ".xltx": "Template without macros",
        ".xltm": "Template with macros",
        ".xlt": "Legacy template format",
        ".xml": "XML spreadsheet format"
    }
    
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("EXTUPDATE - Excel Extension Manager")
        self.window.geometry("800x750")
        self.window.minsize(700, 650)

        try:
            if os.path.exists("assets/extup.ico"):
                self.window.iconbitmap("assets/extup.ico")
        except:
            pass
        
        # Variables
        self.selected_path = tk.StringVar(value="No folder selected...")
        self.current_ext = tk.StringVar(value=".xls")
        self.updated_ext = tk.StringVar(value=".xlsx")
        self.progress_var = tk.DoubleVar()
        self.status_var = tk.StringVar(value="Ready")
        self.backup_var = tk.BooleanVar(value=True)
        self.recursive_var = tk.BooleanVar(value=False)
        
        # Fonts
        self.font_title = Font(family="Arial", size=14, weight="bold")
        self.font_button = Font(family="Arial", size=12, weight="bold")
        self.font_text = Font(family="Arial", size=10)
        
        # History tracking
        self.history_file = "extupdate_history.json"
        self.history = self.load_history()
        
        # Setup UI
        self.setup_ui()
        self.update_button_states()
        
        # Center window
        self.center_window()
        
    def center_window(self):
        """Center the window on screen"""
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f'{width}x{height}+{x}+{y}')
        
    def setup_ui(self):
        """Setup the user interface"""
        # Configure grid weights
        self.window.grid_rowconfigure(1, weight=1)
        self.window.grid_columnconfigure(0, weight=1)
        
        # Create menu
        self.create_menu()
        
        # Header frame
        header_frame = ttk.Frame(self.window, padding="10")
        header_frame.grid(row=0, column=0, sticky="ew")
        
        # Logo/Title
        title_label = ttk.Label(
            header_frame, 
            text="EXTUPDATE", 
            font=Font(family="Arial", size=20, weight="bold")
        )
        title_label.pack()
        
        subtitle_label = ttk.Label(
            header_frame,
            text="Excel Extension Manager",
            font=self.font_text
        )
        subtitle_label.pack()
        
        # Main content frame
        main_frame = ttk.Frame(self.window, padding="10")
        main_frame.grid(row=1, column=0, sticky="nsew")
        main_frame.grid_columnconfigure(1, weight=1)
        
        # Folder selection
        folder_frame = ttk.LabelFrame(main_frame, text="Folder Selection", padding="10")
        folder_frame.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        folder_frame.grid_columnconfigure(1, weight=1)
        
        ttk.Button(
            folder_frame,
            text="Select Folder",
            command=self.select_folder
        ).grid(row=0, column=0, padx=(0, 10))
        
        ttk.Entry(
            folder_frame,
            textvariable=self.selected_path,
            state="readonly",
            font=self.font_text
        ).grid(row=0, column=1, sticky="ew")
        
        # Options frame
        options_frame = ttk.LabelFrame(main_frame, text="Options", padding="10")
        options_frame.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        
        ttk.Checkbutton(
            options_frame,
            text="Create backup before conversion",
            variable=self.backup_var
        ).grid(row=0, column=0, sticky="w", padx=(0, 20))
        
        ttk.Checkbutton(
            options_frame,
            text="Include subfolders (recursive)",
            variable=self.recursive_var
        ).grid(row=0, column=1, sticky="w")
        
        # Extension selection frame
        ext_frame = ttk.LabelFrame(main_frame, text="Extension Conversion", padding="10")
        ext_frame.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        ext_frame.grid_columnconfigure(1, weight=1)
        ext_frame.grid_columnconfigure(3, weight=1)
        
        ttk.Label(ext_frame, text="Current Extension:").grid(row=0, column=0, sticky="e", padx=(0, 5))
        current_combo = ttk.Combobox(
            ext_frame,
            textvariable=self.current_ext,
            values=self.EXCEL_EXTENSIONS,
            state="readonly",
            width=15
        )
        current_combo.grid(row=0, column=1, sticky="w", padx=(0, 20))
        current_combo.bind("<<ComboboxSelected>>", self.on_extension_change)
        
        ttk.Label(ext_frame, text="Convert to:").grid(row=0, column=2, sticky="e", padx=(0, 5))
        updated_combo = ttk.Combobox(
            ext_frame,
            textvariable=self.updated_ext,
            values=self.EXCEL_EXTENSIONS,
            state="readonly",
            width=15
        )
        updated_combo.grid(row=0, column=3, sticky="w")
        updated_combo.bind("<<ComboboxSelected>>", self.on_extension_change)
        
        # Compatibility info
        self.compat_label = ttk.Label(
            ext_frame,
            text="",
            font=Font(family="Arial", size=9, slant="italic"),
            foreground="gray"
        )
        self.compat_label.grid(row=1, column=0, columnspan=4, pady=(5, 0))
        self.update_compatibility_info()
        
        # File preview frame
        preview_frame = ttk.LabelFrame(main_frame, text="Files to be converted", padding="10")
        preview_frame.grid(row=3, column=0, columnspan=3, sticky="nsew", pady=(0, 10))
        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)
        
        # Treeview for file list
        tree_scroll = ttk.Scrollbar(preview_frame)
        tree_scroll.grid(row=0, column=1, sticky="ns")
        
        self.file_tree = ttk.Treeview(
            preview_frame,
            yscrollcommand=tree_scroll.set,
            columns=("size", "modified"),
            height=8
        )
        self.file_tree.grid(row=0, column=0, sticky="nsew")
        tree_scroll.config(command=self.file_tree.yview)
        
        # Configure treeview columns
        self.file_tree.heading("#0", text="File Name")
        self.file_tree.heading("size", text="Size")
        self.file_tree.heading("modified", text="Modified")
        self.file_tree.column("#0", width=400)
        self.file_tree.column("size", width=100)
        self.file_tree.column("modified", width=150)
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(
            main_frame,
            variable=self.progress_var,
            mode='determinate'
        )
        self.progress_bar.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        
        # Button frame with better visibility
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        
        # Create a frame to center the buttons
        button_container = ttk.Frame(button_frame)
        button_container.pack(expand=True)
        
        # Style for buttons to make them more visible
        style = ttk.Style()
        style.configure('Action.TButton', font=self.font_button, padding=(20, 10))
        
        self.update_btn = ttk.Button(
            button_container,
            text="Update Extensions",
            command=self.update_extensions,
            state="disabled",
            style='Action.TButton'
        )
        self.update_btn.pack(side="left", padx=5)
        
        self.clear_btn = ttk.Button(
            button_container,
            text="Clear All",
            command=self.clear_all,
            state="disabled",
            style='Action.TButton'
        )
        self.clear_btn.pack(side="left", padx=5)
        
        ttk.Button(
            button_container,
            text="Exit",
            command=self.window.quit,
            style='Action.TButton'
        ).pack(side="left", padx=5)
        
        # Add instruction label
        instruction_label = ttk.Label(
            main_frame,
            text="Select a folder to enable the Update Extensions button",
            font=Font(family="Arial", size=9, slant="italic"),
            foreground="gray"
        )
        instruction_label.grid(row=6, column=0, columnspan=3, pady=(5, 0))
        
        # Status bar
        status_frame = ttk.Frame(self.window)
        status_frame.grid(row=2, column=0, sticky="ew")
        
        ttk.Label(
            status_frame,
            textvariable=self.status_var,
            font=self.font_text
        ).pack(side="left", padx=10)
        
    def create_menu(self):
        """Create application menu"""
        menubar = tk.Menu(self.window)
        self.window.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Select Folder", command=self.select_folder)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.window.quit)
        
        # Tools menu
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        tools_menu.add_command(label="View History", command=self.show_history)
        tools_menu.add_command(label="Clear History", command=self.clear_history)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="Usage", command=self.show_usage)
        help_menu.add_command(label="About", command=self.show_about)
        
    def select_folder(self):
        """Handle folder selection"""
        folder = filedialog.askdirectory()
        if folder:
            self.selected_path.set(folder)
            self.update_button_states()
            self.scan_files()
            
    def scan_files(self):
        """Scan selected folder for Excel files"""
        self.file_tree.delete(*self.file_tree.get_children())
        
        if not os.path.exists(self.selected_path.get()):
            return
            
        files_found = 0
        current_ext = self.current_ext.get()
        
        if self.recursive_var.get():
            # Recursive search
            for root, dirs, files in os.walk(self.selected_path.get()):
                for file in files:
                    if file.endswith(current_ext):
                        self.add_file_to_tree(os.path.join(root, file))
                        files_found += 1
        else:
            # Non-recursive search
            for file in os.listdir(self.selected_path.get()):
                file_path = os.path.join(self.selected_path.get(), file)
                if os.path.isfile(file_path) and file.endswith(current_ext):
                    self.add_file_to_tree(file_path)
                    files_found += 1
                    
        self.status_var.set(f"Found {files_found} {current_ext} files")
        
    def add_file_to_tree(self, file_path: str):
        """Add file to treeview"""
        stat = os.stat(file_path)
        size = self.format_size(stat.st_size)
        modified = datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M")
        
        # Get relative path if recursive
        if self.recursive_var.get():
            display_name = os.path.relpath(file_path, self.selected_path.get())
        else:
            display_name = os.path.basename(file_path)
            
        self.file_tree.insert("", "end", text=display_name, values=(size, modified))
        
    def format_size(self, size: int) -> str:
        """Format file size in human readable format"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024.0:
                return f"{size:.1f} {unit}"
            size /= 1024.0
        return f"{size:.1f} TB"
        
    def update_extensions(self):
        """Update file extensions with threading"""
        if self.current_ext.get() == self.updated_ext.get():
            messagebox.showwarning("Invalid Selection", "Current and target extensions must be different!")
            return
            
        files = self.get_files_to_convert()
        if not files:
            messagebox.showwarning("No Files", f"No {self.current_ext.get()} files found in the selected directory!")
            return
            
        # Confirm action
        msg = f"Convert {len(files)} files from {self.current_ext.get()} to {self.updated_ext.get()}?"
        if self.backup_var.get():
            msg += "\n\nBackups will be created."
            
        if not messagebox.askyesno("Confirm Conversion", msg):
            return
            
        # Disable buttons during operation
        self.update_btn.config(state="disabled")
        self.clear_btn.config(state="disabled")
        
        # Start conversion in thread
        thread = threading.Thread(target=self.convert_files, args=(files,))
        thread.start()
        
    def get_files_to_convert(self) -> List[str]:
        """Get list of files to convert"""
        files = []
        current_ext = self.current_ext.get()
        
        if self.recursive_var.get():
            for root, dirs, filenames in os.walk(self.selected_path.get()):
                for filename in filenames:
                    if filename.endswith(current_ext):
                        files.append(os.path.join(root, filename))
        else:
            for filename in os.listdir(self.selected_path.get()):
                file_path = os.path.join(self.selected_path.get(), filename)
                if os.path.isfile(file_path) and filename.endswith(current_ext):
                    files.append(file_path)
                    
        return files
        
    def convert_files(self, files: List[str]):
        """Convert files in a separate thread"""
        converted = 0
        errors = []
        
        for i, file_path in enumerate(files):
            try:
                # Update progress
                progress = (i / len(files)) * 100
                self.progress_var.set(progress)
                self.status_var.set(f"Converting: {os.path.basename(file_path)}")
                
                # Create backup if requested
                if self.backup_var.get():
                    backup_path = file_path + ".backup"
                    shutil.copy2(file_path, backup_path)
                
                # Rename file
                new_path = file_path[:-len(self.current_ext.get())] + self.updated_ext.get()
                os.rename(file_path, new_path)
                converted += 1
                
                # Log to history
                self.add_to_history(file_path, new_path)
                
            except Exception as e:
                errors.append(f"{os.path.basename(file_path)}: {str(e)}")
                
        # Update progress to 100%
        self.progress_var.set(100)
        
        # Show results
        self.window.after(0, self.show_conversion_results, converted, errors, len(files))
        
    def show_conversion_results(self, converted: int, errors: List[str], total: int):
        """Show conversion results"""
        self.progress_var.set(0)
        self.status_var.set(f"Conversion complete: {converted}/{total} files")
        
        if errors:
            error_msg = f"Successfully converted {converted} files.\n\nErrors occurred with {len(errors)} files:\n"
            error_msg += "\n".join(errors[:10])  # Show first 10 errors
            if len(errors) > 10:
                error_msg += f"\n... and {len(errors) - 10} more"
            messagebox.showwarning("Conversion Complete with Errors", error_msg)
        else:
            messagebox.showinfo("Success", f"Successfully converted {converted} files!")
            
        # Re-enable buttons and refresh file list
        self.update_btn.config(state="normal")
        self.clear_btn.config(state="normal")
        self.scan_files()
        
    def clear_all(self):
        """Clear all selections"""
        self.selected_path.set("No folder selected...")
        self.file_tree.delete(*self.file_tree.get_children())
        self.status_var.set("Ready")
        self.update_button_states()
        
    def update_button_states(self):
        """Update button states based on current selections"""
        folder_selected = self.selected_path.get() != "No folder selected..."
        
        if folder_selected:
            self.update_btn.config(state="normal")
            self.clear_btn.config(state="normal")
            # Update instruction label if it exists
            for widget in self.window.winfo_children():
                if isinstance(widget, ttk.Frame):
                    for child in widget.winfo_children():
                        if isinstance(child, ttk.Label) and "Select a folder" in child.cget("text"):
                            child.config(text="Ready to convert files", foreground="green")
        else:
            self.update_btn.config(state="disabled")
            self.clear_btn.config(state="disabled")
        
    def on_extension_change(self, event=None):
        """Handle extension change"""
        self.update_compatibility_info()
        if self.selected_path.get() != "No folder selected...":
            self.scan_files()
            
    def update_compatibility_info(self):
        """Update compatibility information label"""
        target_ext = self.updated_ext.get()
        info = self.COMPATIBILITY_WARNINGS.get(target_ext, "")
        self.compat_label.config(text=info)
        
    def add_to_history(self, old_path: str, new_path: str):
        """Add conversion to history"""
        entry = {
            "timestamp": datetime.now().isoformat(),
            "old_path": old_path,
            "new_path": new_path,
            "old_ext": self.current_ext.get(),
            "new_ext": self.updated_ext.get()
        }
        self.history.append(entry)
        self.save_history()
        
    def load_history(self) -> List[Dict]:
        """Load conversion history"""
        if os.path.exists(self.history_file):
            try:
                with open(self.history_file, 'r') as f:
                    return json.load(f)
            except:
                return []
        return []
        
    def save_history(self):
        """Save conversion history"""
        try:
            with open(self.history_file, 'w') as f:
                json.dump(self.history[-1000:], f, indent=2)  # Keep last 1000 entries
        except:
            pass
            
    def show_history(self):
        """Show conversion history window"""
        history_window = tk.Toplevel(self.window)
        history_window.title("Conversion History")
        history_window.geometry("600x400")
        
        # Create treeview
        tree_frame = ttk.Frame(history_window, padding="10")
        tree_frame.pack(fill="both", expand=True)
        
        tree_scroll = ttk.Scrollbar(tree_frame)
        tree_scroll.pack(side="right", fill="y")
        
        history_tree = ttk.Treeview(
            tree_frame,
            yscrollcommand=tree_scroll.set,
            columns=("old", "new", "time"),
            height=15
        )
        history_tree.pack(fill="both", expand=True)
        tree_scroll.config(command=history_tree.yview)
        
        # Configure columns
        history_tree.heading("#0", text="ID")
        history_tree.heading("old", text="Original")
        history_tree.heading("new", text="Converted")
        history_tree.heading("time", text="Date/Time")
        
        history_tree.column("#0", width=50)
        history_tree.column("old", width=200)
        history_tree.column("new", width=200)
        history_tree.column("time", width=130)
        
        # Add history entries
        for i, entry in enumerate(reversed(self.history[-100:])):  # Show last 100
            timestamp = datetime.fromisoformat(entry["timestamp"]).strftime("%Y-%m-%d %H:%M")
            history_tree.insert(
                "", "end",
                text=str(i+1),
                values=(
                    os.path.basename(entry["old_path"]),
                    os.path.basename(entry["new_path"]),
                    timestamp
                )
            )
            
    def clear_history(self):
        """Clear conversion history"""
        if messagebox.askyesno("Clear History", "Are you sure you want to clear the conversion history?"):
            self.history = []
            self.save_history()
            messagebox.showinfo("History Cleared", "Conversion history has been cleared.")
            
    def show_usage(self):
        """Show usage information"""
        usage_text = """How to use EXTUPDATE:

        1. Click 'Select Folder' to choose a directory containing Excel files
        2. Select the current file extension from the dropdown
        3. Select the target extension for conversion
        4. Optional: Enable backup creation and/or recursive search
        5. Review the files that will be converted in the preview
        6. Click 'Update Extensions' to start the conversion

        Tips:
        • Always create backups when converting between incompatible formats
        • Use recursive search to convert files in subfolders
        • Check the compatibility info before converting
        • View conversion history from the Tools menu"""
        
        messagebox.showinfo("Usage Instructions", usage_text)
        
    def show_about(self):
        """Show about information"""
        about_text = """EXTUPDATE - Excel Extension Manager

        Version: 2.0.0
        Author: sorzkode

        An tool for bulk conversion of Excel file extensions.

        Features:
        • Batch file extension conversion
        • Backup creation option
        • Recursive folder processing
        • Conversion history tracking
        • File preview with metadata
        • Compatibility warnings

        License: MIT
        GitHub: https://github.com/sorzkode/extupdate"""
        
        messagebox.showinfo("About EXTUPDATE", about_text)
        
    def run(self):
        """Run the application"""
        self.window.mainloop()


def main():
    """Main entry point"""
    app = ExtUpdateApp()
    app.run()


if __name__ == "__main__":
    main()