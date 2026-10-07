"""
Run tab (Standard Import)
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import platform
from pathlib import Path
from typing import TYPE_CHECKING

from gui.models import FileError

if TYPE_CHECKING:
    from gui.meeting_import_app import MeetingImportGUI

# ================================================================
# TAB 5: RUN
# ================================================================

def build_run_tab(app):
    """Build the run/preview tab."""
    tab = app.tab_frames["Run"]
    
    # Header with prominent action buttons
    header = ctk.CTkFrame(tab, fg_color="transparent")
    header.pack(fill="x", padx=10, pady=(10, 5))
    
    ctk.CTkLabel(header, text="Import Preview & Execution",
                font=ctk.CTkFont(size=16, weight="bold"),
                anchor="w").pack(side="left")
    
    # Right side - prominent action buttons and status
    right_header = ctk.CTkFrame(header, fg_color="transparent")
    right_header.pack(side="right")
    
    # Dry Run button (prominent, large)
    app.dry_run_btn = ctk.CTkButton(right_header, text="🔍 Dry Run (Preview)",
                                    width=180, height=40,
                                    font=ctk.CTkFont(size=14, weight="bold"),
                                    fg_color="#555555",
                                    hover_color="#666666",
                                    command=app._run_dry_run)
    app.dry_run_btn.pack(side="left", padx=5)
    
    # Run Import button (prominent, large, green)
    app.run_import_btn = ctk.CTkButton(right_header, text="▶ Run Import",
                                       width=150, height=40,
                                       font=ctk.CTkFont(size=14, weight="bold"),
                                       fg_color="green",
                                       hover_color="darkgreen",
                                       command=app._run_import)
    app.run_import_btn.pack(side="left", padx=5)
    
    # Status label
    app.status_label = ctk.CTkLabel(right_header, text="No results",
                                    text_color="gray", 
                                    font=ctk.CTkFont(size=12),
                                    anchor="e")
    app.status_label.pack(side="left", padx=(10, 0))
    
    # Preview treeview
    tree_frame = ctk.CTkFrame(tab)
    tree_frame.pack(fill="both", expand=True, padx=10, pady=(5, 10))
    
    # Create treeview
    columns = ("event_name", "event_date", "event_time", "meeting_type",
              "event_category", "video", "agenda")
    app.preview_tree = ttk.Treeview(tree_frame, columns=columns,
                                     show="headings", height=15,
                                     style="Dark.Treeview")
    
    # Configure columns
    app.preview_tree.heading("event_name", text="Event Name")
    app.preview_tree.heading("event_date", text="Date")
    app.preview_tree.heading("event_time", text="Time")
    app.preview_tree.heading("meeting_type", text="Meeting Type")
    app.preview_tree.heading("event_category", text="Event Category")
    app.preview_tree.heading("video", text="Video")
    app.preview_tree.heading("agenda", text="Agenda")
    
    app.preview_tree.column("event_name", width=250)
    app.preview_tree.column("event_date", width=100)
    app.preview_tree.column("event_time", width=100)
    app.preview_tree.column("meeting_type", width=150)
    app.preview_tree.column("event_category", width=150)
    app.preview_tree.column("video", width=100)
    app.preview_tree.column("agenda", width=100)
    
    # Scrollbars
    scrollbar_y = ttk.Scrollbar(tree_frame, orient="vertical",
                               command=app.preview_tree.yview)
    scrollbar_x = ttk.Scrollbar(tree_frame, orient="horizontal",
                               command=app.preview_tree.xview)
    app.preview_tree.configure(yscrollcommand=scrollbar_y.set,
                               xscrollcommand=scrollbar_x.set)
    
    app.preview_tree.grid(row=0, column=0, sticky="nsew")
    scrollbar_y.grid(row=0, column=1, sticky="ns")
    scrollbar_x.grid(row=1, column=0, sticky="ew")
    
    tree_frame.grid_rowconfigure(0, weight=1)
    tree_frame.grid_columnconfigure(0, weight=1)
    
    # Style treeview for dark theme using a named style (isolated from other tabs)
    style = ttk.Style()
    style.theme_use("default")
    style.configure("Dark.Treeview",
                   background="#2b2b2b",
                   foreground="white",
                   fieldbackground="#2b2b2b",
                   borderwidth=0)
    style.configure("Dark.Treeview.Heading",
                   background="#1f1f1f",
                   foreground="white",
                   borderwidth=1)
    style.map("Dark.Treeview", background=[("selected", "#144870")])
    
    # Messages tabview
    msg_tabview = ctk.CTkTabview(tab)
    msg_tabview.pack(fill="x", padx=10, pady=(0, 10))
    
    msg_tabview.add("Warnings")
    msg_tabview.add("Duplicates")
    
    app.warnings_text = ctk.CTkTextbox(msg_tabview.tab("Warnings"), height=100)
    app.warnings_text.pack(fill="both", expand=True, padx=5, pady=5)
    
    app.duplicates_text = ctk.CTkTextbox(msg_tabview.tab("Duplicates"), height=100)
    app.duplicates_text.pack(fill="both", expand=True, padx=5, pady=5)
    
    # Action bar - bottom controls
    action_bar = ctk.CTkFrame(tab)
    action_bar.pack(fill="x", padx=10, pady=(0, 10))
    
    # Right side - export button only
    right_frame = ctk.CTkFrame(action_bar, fg_color="transparent")
    right_frame.pack(side="right", padx=10, pady=10)
    
    app.export_btn = ctk.CTkButton(right_frame, text="📊 Export Excel",
                                   width=140, height=36,
                                   command=app._export_excel,
                                   state="disabled")
    app.export_btn.pack(side="left", padx=5)

def run_dry_run(app):
    """Execute dry run."""
    app._execute_import(dry_run=True)

def run_import(app):
    """Execute actual import."""
    response = messagebox.askyesno("Confirm Import",
                                  "This will perform the actual import.\nContinue?")
    if response:
        app._execute_import(dry_run=False)

def execute_import(app, dry_run=True):
    """Execute the import process."""
    # Validate configuration
    if not app.base_dir.get():
        messagebox.showerror("Validation Error", "Please select a base directory.")
        return
    
    if not app.meeting_bodies:
        messagebox.showerror("Validation Error", "Please add at least one meeting body.")
        return
    
    # Build config
    config = {
        "base_dir": app.base_dir.get(),
        "date_format": app.date_format.get(),
        "dry_run": dry_run,
        "meeting_bodies": [body.to_dict() for body in app.meeting_bodies],
        "subtypes": [st.to_dict() for st in app.subtypes]
    }
    
    # Show loading
    app.dry_run_btn.configure(state="disabled")
    app.run_import_btn.configure(state="disabled")
    app.status_label.configure(text="Running...", text_color="yellow")
    app.update()
    
    # Clear previous errors (fresh start for each run)
    app.file_errors.clear()
    app._refresh_file_errors_display()
    
    try:
        # Call backend
        if app.backend_wrapper:
            result = app.backend_wrapper.run_import(config)
        else:
            # Fallback mock
            result = app._mock_import(config)
        
        if result:
            app.last_result = result
            app._display_results(result)
            
            # Populate errors tab with NEW errors from this run
            if result.get("errors"):
                app._populate_file_errors(result["errors"], dry_run=dry_run)
            
            # Only enable export for actual run (not dry run)
            if not dry_run:
                app.export_btn.configure(state="normal")
            else:
                app.export_btn.configure(state="disabled")
            
            # Update status
            record_count = len(result.get("records", []))
            error_count = len(result.get("errors", []))
            warning_count = len(result.get("warnings", []))
            
            status_text = f"{'[DRY RUN] ' if dry_run else ''}{record_count} records"
            if error_count:
                status_text += f", {error_count} errors"
            if warning_count:
                status_text += f", {warning_count} warnings"
            
            status_color = "red" if error_count else ("yellow" if warning_count else "green")
            app.status_label.configure(text=status_text, text_color=status_color)
    
    except Exception as e:
        messagebox.showerror("Import Error", f"Failed to run import:\n{str(e)}")
        app.status_label.configure(text="Error", text_color="red")
    
    finally:
        app.dry_run_btn.configure(state="normal")
        app.run_import_btn.configure(state="normal")

def mock_import(app, config):
    """Mock import for testing without backend."""
    return {
        "records": [
            {
                "Event Name": "City Council Meeting",
                "Event Date": "01/15/2024",
                "Event Time": "06:00:00 PM",
                "Meeting Type": "City Council",
                "Event Category": "City Council",
                "Video File Name": "Council_01_15_2024.mp4",
                "Agenda File Name": "Agenda_01_15_2024.pdf"
            }
        ],
        "warnings": ["Sample warning message"],
        "duplicates": [],
        "errors": []
    }

def display_results(app, result: dict):
    """Display import results."""
    # Clear tree
    for item in app.preview_tree.get_children():
        app.preview_tree.delete(item)
    
    # Render missing values (None / pandas NaN) as blank, never "nan".
    def _cell(v):
        if v is None:
            return ""
        s = str(v)
        return "" if s.strip().lower() == "nan" else s

    # Populate tree
    for record in result.get("records", []):
        values = (
            _cell(record.get("Event Name", "")),
            _cell(record.get("Event Date", "")),
            _cell(record.get("Event Time", "")),
            _cell(record.get("Meeting Type", "")),
            _cell(record.get("Event Category", "")),
            "✓" if record.get("Video File Name") else "",
            "✓" if record.get("Agenda File Name") else ""
        )
        app.preview_tree.insert("", "end", values=values)
    
    # Display warnings
    app.warnings_text.delete("1.0", "end")
    if result.get("warnings"):
        app.warnings_text.insert("1.0", "\n".join(result["warnings"]))
    else:
        app.warnings_text.insert("1.0", "No warnings")
    
    # Display duplicates
    app.duplicates_text.delete("1.0", "end")
    if result.get("duplicates"):
        app.duplicates_text.insert("1.0", "\n".join(result["duplicates"]))
    else:
        app.duplicates_text.insert("1.0", "No duplicates")

def populate_file_errors(app, errors: list, dry_run: bool = False):
    """Populate file errors tab from import results."""
    # Errors were already cleared in _execute_import
    
    for error_data in errors:
        error = FileError(
            error_type=error_data.get("error_type", "Unknown"),
            file_name=error_data.get("file_name", ""),
            issue=error_data.get("issue", ""),
            original_location=error_data.get("original_location", ""),
            current_location=error_data.get("current_location", ""),
            file_path=error_data.get("file_path", "")
        )
        app.file_errors.append(error)
    
    app._refresh_file_errors_display()
    
    # Switch to errors tab if there are errors
    if app.file_errors:
        app._switch_tab("File Errors")
        
        # Show different message for dry run vs actual run
        if dry_run:
            messagebox.showinfo("Errors Found (Dry Run)", 
                              f"Found {len(app.file_errors)} errors in preview.\n\n" +
                              "Fix these errors before running actual import:\n" +
                              "• Ambiguous/Weekend Dates: Rename files\n" +
                              "• Duplicates: Use 'Remove' button\n" +
                              "• Unrecognized: Already moved to Review folder")
        else:
            messagebox.showwarning("Errors Found", 
                                 f"Found {len(app.file_errors)} errors during import.\n\n" +
                                 "Please review and fix in Tab 4 (File Errors).")

def open_review_folders(app):
    """Open the review folders."""
    base = app.base_dir.get()
    if not base:
        messagebox.showwarning("No Base Directory", "Please select a base directory first.")
        return
    
    review_path = Path(base) / "Review_Files"
    if review_path.exists():
        import platform
        if platform.system() == "Windows":
            os.startfile(str(review_path))
        elif platform.system() == "Darwin":
            os.system(f'open "{review_path}"')
        else:
            os.system(f'xdg-open "{review_path}"')
    else:
        messagebox.showinfo("No Review Folder", "Review folder will be created after first import.")

def export_excel(app):
    """Export results to Excel as Final_Import_File.xlsx."""
    if not app.last_result:
        messagebox.showwarning("No Results", "No results to export.")
        return
    
    # Default filename
    default_name = "Final_Import_File.xlsx"
    
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        initialfile=default_name,
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        title="Export Final Import File"
    )
    
    if file_path:
        try:
            import pandas as pd
            df = pd.DataFrame(app.last_result["records"])
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Success", 
                              f"Exported as:\n{Path(file_path).name}\n\n" +
                              "Upload this file to CivicClerk import page.")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export:\n{str(e)}")

