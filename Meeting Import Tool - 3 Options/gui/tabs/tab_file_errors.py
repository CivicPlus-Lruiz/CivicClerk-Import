"""
File Errors tab (Standard Import)
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import shutil
from pathlib import Path
from typing import TYPE_CHECKING, List
from gui.models import FileError
if TYPE_CHECKING:
    from gui.meeting_import_app import MeetingImportGUI

# ================================================================
# TAB 4: FILE ERRORS
# ================================================================

def build_file_errors_tab(app):
    """Build the file errors resolution tab."""
    tab = app.tab_frames["File Errors"]
    
    # Header
    header = ctk.CTkFrame(tab)
    header.pack(fill="x", padx=10, pady=10)
    
    ctk.CTkLabel(header, text="File Errors - Interactive Resolution", 
                font=ctk.CTkFont(size=16, weight="bold")).pack(side="left", padx=5)
    
    app.error_count_label = ctk.CTkLabel(header, text="No errors",
                                         text_color="gray")
    app.error_count_label.pack(side="right", padx=10)
    
    # Action buttons
    button_frame = ctk.CTkFrame(tab)
    button_frame.pack(fill="x", padx=10, pady=5)
    
    ctk.CTkButton(button_frame, text="Rename Selected", width=140,
                 command=app._rename_selected_errors).pack(side="left", padx=5)
    
    ctk.CTkButton(button_frame, text="🔍 Run OCR on Selected", width=180,
                 fg_color="#555555", hover_color="#666666",
                 command=app._run_ocr_on_selected).pack(side="left", padx=5)

    ctk.CTkButton(button_frame, text="📄 Preview PDF", width=140,
                 fg_color="#1a5276", hover_color="#154360",
                 command=app._preview_selected_pdf).pack(side="left", padx=5)

    ctk.CTkButton(button_frame, text="Remove (Duplicates)", width=160,
                 fg_color="#c44536", hover_color="#a03529",
                 command=app._remove_duplicates).pack(side="left", padx=5)
    
    ctk.CTkButton(button_frame, text="📂 Open Review Folders", width=180,
                 command=app._open_review_folders).pack(side="left", padx=5)
    
    ctk.CTkButton(button_frame, text="Open Location", width=140,
                 command=app._open_error_location).pack(side="left", padx=5)
    
    ctk.CTkButton(button_frame, text="Mark Resolved", width=140,
                 command=app._mark_errors_resolved).pack(side="left", padx=5)
    
    ctk.CTkButton(button_frame, text="Re-run Dry Run", width=140,
                 fg_color="green", hover_color="darkgreen",
                 command=app._run_dry_run).pack(side="left", padx=5)
    
    ctk.CTkButton(button_frame, text="Select All", width=100,
                 command=app._select_all_errors).pack(side="left", padx=5)
    
    ctk.CTkButton(button_frame, text="Deselect All", width=100,
                 command=app._deselect_all_errors).pack(side="left", padx=5)

    # Table headers
    headers_frame = ctk.CTkFrame(tab)
    headers_frame.pack(fill="x", padx=10, pady=5)
    
    ctk.CTkLabel(headers_frame, text="☐", width=30,
                font=ctk.CTkFont(weight="bold"), anchor="w").pack(side="left", padx=2)
    
    ctk.CTkLabel(headers_frame, text="Error Type", width=120,
                font=ctk.CTkFont(weight="bold"), anchor="w").pack(side="left", padx=2)
    
    ctk.CTkLabel(headers_frame, text="File Type", width=120,
                font=ctk.CTkFont(weight="bold"), anchor="w").pack(side="left", padx=2)
    
    ctk.CTkLabel(headers_frame, text="File Name", width=250,
                font=ctk.CTkFont(weight="bold"), anchor="w").pack(side="left", padx=2)
    
    ctk.CTkLabel(headers_frame, text="Issue", width=250,
                font=ctk.CTkFont(weight="bold"), anchor="w").pack(side="left", padx=2)
    
    ctk.CTkLabel(headers_frame, text="Suggested Filename", width=280,
                font=ctk.CTkFont(weight="bold"), anchor="w").pack(side="left", padx=2)
    
    ctk.CTkLabel(headers_frame, text="Action", width=140,
                font=ctk.CTkFont(weight="bold"), anchor="w").pack(side="left", padx=2)
    
    # Scrollable errors list
    app.errors_scroll = ctk.CTkScrollableFrame(tab)
    app.errors_scroll.pack(fill="both", expand=True, padx=10, pady=10)
    
    app._refresh_file_errors_display()

def refresh_file_errors_display(app):
    """Rebuild the file errors list display."""
    # Clear existing
    for widget in app.errors_scroll.winfo_children():
        widget.destroy()
    
    if not app.file_errors:
        label = ctk.CTkLabel(app.errors_scroll, text="No errors to display",
                            text_color="gray", font=ctk.CTkFont(size=14))
        label.pack(pady=50)
        app.error_count_label.configure(text="No errors")
        return
    
    # Update count
    app.error_count_label.configure(text=f"{len(app.file_errors)} errors",
                                    text_color="red")
    
    # Create row for each error
    for idx, error in enumerate(app.file_errors):
        # Alternating row colors
        row_color = "#2b2b2b" if idx % 2 == 0 else "#333333"
        
        row = ctk.CTkFrame(app.errors_scroll, fg_color=row_color)
        row.pack(fill="x", pady=2)
        
        # Checkbox
        var = ctk.BooleanVar(value=error.selected)
        checkbox = ctk.CTkCheckBox(row, text="", variable=var, width=30,
                                   command=lambda i=idx, v=var: setattr(app.file_errors[i], 'selected', v.get()))
        checkbox.pack(side="left", padx=2)
        
        # Error Type (fixed width)
        ctk.CTkLabel(row, text=error.error_type, width=120, anchor="w").pack(side="left", padx=2)
        
        # File Type - parsed from file_path (folder two levels up from filename)
        file_type = ""
        if hasattr(error, 'file_path') and error.file_path:
            try:
                parts = Path(error.file_path).parts
                # Structure: .../meetings/<Body>/<FileType>/filename
                # So FileType is parts[-2]
                if len(parts) >= 2:
                    file_type = parts[-2]
            except:
                pass
        ctk.CTkLabel(row, text=file_type or "—", width=120, anchor="w").pack(side="left", padx=2)
        
        # File Name
        ctk.CTkLabel(row, text=error.file_name, width=250, anchor="w", 
                    wraplength=240).pack(side="left", padx=2)
        
        # Issue
        issue_text = error.issue
        # Strip [DRY RUN] prefix if present before pattern matching
        dry_run_prefix = "[DRY RUN] "
        clean_issue = issue_text[len(dry_run_prefix):] if issue_text.startswith(dry_run_prefix) else issue_text
        if clean_issue.lower().startswith("duplicate file detected - kept:"):
            issue_text = "Kept: " + clean_issue.split("kept:", 1)[-1].strip()
        ctk.CTkLabel(row, text=issue_text, width=250, anchor="w",
                    wraplength=240).pack(side="left", padx=2)
        
        # OCR Suggestion column (only for ambiguous dates)
        if error.error_type.lower() == "ambiguous date":
            if hasattr(error, 'ocr_suggestion') and error.ocr_suggestion:
                # Show suggestion with confidence
                suggestion_text = error.ocr_suggestion
                if hasattr(error, 'ocr_confidence') and error.ocr_confidence:
                    suggestion_text += f" ({error.ocr_confidence})"
                
                suggestion_label = ctk.CTkLabel(row, text=suggestion_text, width=280, anchor="w",
                                               text_color="#7eb8f7", wraplength=270)
                suggestion_label.pack(side="left", padx=2)
                
                # Action buttons
                if error.ocr_suggestion != "OCR Unable to parse data":
                    action_frame = ctk.CTkFrame(row, fg_color="transparent", width=140)
                    action_frame.pack(side="left", padx=2)
                    
                    # Apply button
                    ctk.CTkButton(
                        action_frame,
                        text="✓ Apply",
                        width=60,
                        height=24,
                        fg_color="green",
                        hover_color="darkgreen",
                        command=lambda i=idx: app._apply_ocr_suggestion(i)
                    ).pack(side="left", padx=2)
                    
                    # Edit button
                    ctk.CTkButton(
                        action_frame,
                        text="✏️ Edit",
                        width=60,
                        height=24,
                        command=lambda i=idx: app._edit_ocr_suggestion(i)
                    ).pack(side="left", padx=2)
                else:
                    # Just show placeholder
                    ctk.CTkLabel(row, text="—", width=140, anchor="w").pack(side="left", padx=2)
            else:
                # No suggestion yet
                ctk.CTkLabel(row, text="(Run OCR to get suggestion)", width=280, 
                            anchor="w", text_color="gray").pack(side="left", padx=2)
                ctk.CTkLabel(row, text="", width=140).pack(side="left", padx=2)
        elif error.error_type.lower() == "duplicate":
            # Duplicate rows: show empty suggestion column + Compare button
            ctk.CTkLabel(row, text="—", width=280, anchor="w").pack(side="left", padx=2)
            ctk.CTkButton(
                row,
                text="🔍 Compare",
                width=120,
                height=24,
                fg_color="#1a5276",
                hover_color="#154360",
                command=lambda i=idx: app._compare_duplicate(i)
            ).pack(side="left", padx=2)
        else:
            # All other error types - show empty columns
            ctk.CTkLabel(row, text="—", width=280, anchor="w").pack(side="left", padx=2)
            ctk.CTkLabel(row, text="", width=140).pack(side="left", padx=2)

def run_ocr_on_selected(app):
    """Run OCR on selected ambiguous date files."""
    import importlib.util
    import sys
    import threading

    # Load OCRHandler dynamically from project root (same folder as main.py)
    _candidates = [
        os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "ocr_handler.py"),
        os.path.join(os.path.dirname(sys.argv[0]), "ocr_handler.py"),
        os.path.join(os.getcwd(), "ocr_handler.py"),
    ]
    _ocr_path = next((p for p in _candidates if os.path.exists(p)), None)
    if _ocr_path is None:
        messagebox.showerror("OCR Not Found",
            "Could not find ocr_handler.py.\n\nMake sure it is in the same folder as main.py.")
        return
    _spec = importlib.util.spec_from_file_location("ocr_handler", _ocr_path)
    _mod = importlib.util.module_from_spec(_spec)
    _spec.loader.exec_module(_mod)
    OCRHandler = _mod.OCRHandler

    # Get selected ambiguous date errors
    ambiguous_errors = [e for e in app.file_errors 
                       if e.selected and e.error_type.lower() == "ambiguous date"]
    
    if not ambiguous_errors:
        messagebox.showinfo("No Selection", 
            "Please select files with ambiguous dates to process.")
        return
    
    # Confirm
    response = messagebox.askyesno(
        "Run OCR",
        f"Run OCR analysis on {len(ambiguous_errors)} file(s)?\n\n"
        "This may take a few moments depending on file size."
    )
    
    if not response:
        return
    
    # Create progress dialog
    progress_window = ctk.CTkToplevel(app)
    progress_window.title("OCR Processing")
    progress_window.geometry("500x200")
    progress_window.transient(app)
    progress_window.grab_set()
    
    # Progress label
    progress_label = ctk.CTkLabel(
        progress_window,
        text="Initializing OCR...",
        font=ctk.CTkFont(size=14)
    )
    progress_label.pack(pady=20)
    
    # Progress bar
    progress_bar = ctk.CTkProgressBar(progress_window, width=400)
    progress_bar.pack(pady=10)
    progress_bar.set(0)
    
    # Current file label
    current_file_label = ctk.CTkLabel(
        progress_window,
        text="",
        font=ctk.CTkFont(size=11),
        text_color="gray"
    )
    current_file_label.pack(pady=5)
    
    # Cancel button
    cancel_flag = {"cancelled": False}
    
    def cancel_ocr():
        cancel_flag["cancelled"] = True
        cancel_btn.configure(state="disabled", text="Cancelling...")
    
    cancel_btn = ctk.CTkButton(
        progress_window,
        text="Cancel",
        command=cancel_ocr,
        fg_color="#c44536",
        hover_color="#a03529"
    )
    cancel_btn.pack(pady=10)
    
    # Run OCR in thread
    def process_ocr():
        handler = OCRHandler()
        
        for idx, error in enumerate(ambiguous_errors):
            if cancel_flag["cancelled"]:
                break
            
            # Update progress
            progress = (idx + 1) / len(ambiguous_errors)
            app.after(0, lambda p=progress: progress_bar.set(p))
            app.after(0, lambda i=idx: progress_label.configure(
                text=f"Processing {i+1} of {len(ambiguous_errors)}..."
            ))
            app.after(0, lambda f=error.file_name: current_file_label.configure(
                text=f"Current: {f}"
            ))
            
            # Extract meeting body and file type from path
            file_path = error.file_path if hasattr(error, 'file_path') else ""
            meeting_body = ""
            file_type = ""
            
            if file_path:
                # Extract from path structure: .../meetings/MeetingBody/FileType/file.pdf
                parts = file_path.replace("\\", "/").split("/")
                if "meetings" in parts:
                    meetings_idx = parts.index("meetings")
                    if len(parts) > meetings_idx + 2:
                        meeting_body = parts[meetings_idx + 1]
                        file_type = parts[meetings_idx + 2]
            
            # Run OCR
            result = handler.suggest_filename(
                file_path,
                meeting_body=meeting_body,
                file_type=file_type
            )
            
            # Store result in error object
            error.ocr_suggestion = result.get('suggested_name')
            error.ocr_confidence = result.get('confidence')
            error.ocr_date = result.get('date')
            error.ocr_source = result.get('source')
        
        # Close dialog and refresh display
        app.after(0, progress_window.destroy)
        app.after(0, app._refresh_file_errors_display)
        
        if not cancel_flag["cancelled"]:
            app.after(0, lambda: messagebox.showinfo(
                "OCR Complete",
                f"Processed {len(ambiguous_errors)} file(s).\n\n"
                "Review suggestions in the 'Suggested Filename' column."
            ))
    
    # Start OCR thread
    thread = threading.Thread(target=process_ocr, daemon=True)
    thread.start()

def apply_ocr_suggestion(app, error_index):
    """Apply OCR suggested filename to a file."""
    error = app.file_errors[error_index]
    
    if not hasattr(error, 'ocr_suggestion') or not error.ocr_suggestion:
        messagebox.showwarning("No Suggestion", "No OCR suggestion available for this file.")
        return
    
    if error.ocr_suggestion == "OCR Unable to parse data":
        messagebox.showwarning("Unable to Parse", 
            "OCR was unable to determine a suggested filename for this file.")
        return
    
    # Get old and new paths
    old_path = error.file_path if hasattr(error, 'file_path') else ""
    if not old_path or not os.path.exists(old_path):
        messagebox.showerror("File Not Found", f"Cannot find file: {error.file_name}")
        return
    
    # Build new path
    new_path = os.path.join(os.path.dirname(old_path), error.ocr_suggestion)
    
    # Check if file already exists
    if os.path.exists(new_path):
        messagebox.showerror("File Exists", 
            f"A file with the suggested name already exists:\n{error.ocr_suggestion}")
        return
    
    # Confirm rename
    response = messagebox.askyesno(
        "Confirm Rename",
        f"Rename file?\n\n"
        f"From: {error.file_name}\n"
        f"To: {error.ocr_suggestion}\n\n"
        f"Confidence: {error.ocr_confidence or 'N/A'}"
    )
    
    if not response:
        return
    
    # Perform rename
    try:
        os.rename(old_path, new_path)
        
        # Remove from errors list
        app.file_errors.pop(error_index)
        app._refresh_file_errors_display()
        
        messagebox.showinfo("Success", f"File renamed successfully!")
        
    except Exception as e:
        messagebox.showerror("Rename Failed", f"Failed to rename file:\n\n{str(e)}")

def edit_ocr_suggestion(app, error_index):
    """Allow user to edit OCR suggestion before applying."""
    error = app.file_errors[error_index]
    
    # Create edit dialog
    dialog = ctk.CTkToplevel(self)
    dialog.title("Edit Suggested Filename")
    dialog.geometry("600x250")
    dialog.transient(self)
    dialog.grab_set()
    
    # Original filename
    ctk.CTkLabel(
        dialog,
        text=f"Original: {error.file_name}",
        font=ctk.CTkFont(size=12)
    ).pack(pady=10, padx=20, anchor="w")
    
    # Suggested filename (editable)
    ctk.CTkLabel(
        dialog,
        text="Suggested Filename:",
        font=ctk.CTkFont(size=12, weight="bold")
    ).pack(pady=(10, 5), padx=20, anchor="w")
    
    suggestion_entry = ctk.CTkEntry(dialog, width=550)
    suggestion_entry.pack(pady=5, padx=20)
    suggestion_entry.insert(0, error.ocr_suggestion or "")
    
    # Confidence
    if hasattr(error, 'ocr_confidence') and error.ocr_confidence:
        ctk.CTkLabel(
            dialog,
            text=f"Confidence: {error.ocr_confidence}",
            font=ctk.CTkFont(size=11),
            text_color="gray"
        ).pack(pady=5)
    
    # Buttons
    button_frame = ctk.CTkFrame(dialog, fg_color="transparent")
    button_frame.pack(pady=20)
    
    def apply_edited():
        new_name = suggestion_entry.get().strip()
        if not new_name:
            messagebox.showwarning("Empty Name", "Please enter a filename.")
            return
        
        if not new_name.endswith('.pdf'):
            new_name += '.pdf'
        
        error.ocr_suggestion = new_name
        dialog.destroy()
        app._apply_ocr_suggestion(error_index)
    
    ctk.CTkButton(
        button_frame,
        text="Apply",
        command=apply_edited,
        fg_color="green",
        hover_color="darkgreen",
        width=120
    ).pack(side="left", padx=5)
    
    ctk.CTkButton(
        button_frame,
        text="Cancel",
        command=dialog.destroy,
        width=120
    ).pack(side="left", padx=5)

def rename_selected_errors(app):
    """Rename selected error files."""
    selected = [e for e in app.file_errors if e.selected]
    if not selected:
        messagebox.showwarning("No Selection", "Please select files to rename.")
        return
    
    if len(selected) == 1:
        # Single rename
        error = selected[0]
        app._rename_single_file(error)
    else:
        # Bulk rename
        app._bulk_rename_files(selected)

def rename_single_file(app, error: FileError):
    """Rename a single file."""
    # Create rename dialog
    dialog = ctk.CTkToplevel(app)
    dialog.title("Rename File")
    dialog.geometry("600x200")
    dialog.transient(app)
    dialog.grab_set()
    
    ctk.CTkLabel(dialog, text=f"Current name: {error.file_name}",
                font=ctk.CTkFont(size=12)).pack(pady=10)
    
    ctk.CTkLabel(dialog, text="New name:").pack(pady=5)
    
    new_name_var = ctk.StringVar(value=error.file_name)
    entry = ctk.CTkEntry(dialog, textvariable=new_name_var, width=500)
    entry.pack(pady=5)
    
    def do_rename():
        new_name = new_name_var.get().strip()
        if not new_name:
            messagebox.showerror("Invalid Name", "Please enter a valid filename.")
            return
        
        try:
            # Use the file_path directly (full path to file)
            old_path = Path(error.file_path)
            new_path = old_path.parent / new_name
            
            if old_path.exists():
                old_path.rename(new_path)
                
                # Copy renamed file to Renamed Files folder in base directory
                renamed_folder = Path(app.base_dir.get()) / "Renamed Files"
                renamed_folder.mkdir(exist_ok=True)
                shutil.copy2(str(new_path), str(renamed_folder / new_name))
                
                # Remove from error list
                app.file_errors.remove(error)
                app._refresh_file_errors_display()
                
                dialog.destroy()
                messagebox.showinfo("Success", f"File renamed to:\n{new_name}\n\nCopy added to:\nRenamed Files/{new_name}")
            else:
                messagebox.showerror("File Not Found", f"Could not find file:\n{old_path}")
        
        except Exception as e:
            messagebox.showerror("Rename Failed", f"Error: {str(e)}")
    
    button_frame = ctk.CTkFrame(dialog)
    button_frame.pack(pady=20)
    
    ctk.CTkButton(button_frame, text="Rename", width=150,
                 command=do_rename).pack(side="left", padx=5)
    ctk.CTkButton(button_frame, text="Cancel", width=100,
                 command=dialog.destroy).pack(side="left", padx=5)

def bulk_rename_files(app, errors: List[FileError]):
    """Bulk rename multiple files."""
    # Create bulk rename dialog
    dialog = ctk.CTkToplevel(self)
    dialog.title("Bulk Rename Files")
    dialog.geometry("700x400")
    dialog.transient(self)
    dialog.grab_set()
    
    ctk.CTkLabel(dialog, text=f"Bulk Rename {len(errors)} Files",
                font=ctk.CTkFont(size=16, weight="bold")).pack(pady=10)
    
    # Pattern options
    pattern_frame = ctk.CTkFrame(dialog)
    pattern_frame.pack(fill="x", padx=20, pady=10)
    
    ctk.CTkLabel(pattern_frame, text="Rename Pattern:").pack(anchor="w")
    
    pattern_var = ctk.StringVar(value="prefix")
    ctk.CTkRadioButton(pattern_frame, text="Add Prefix", variable=pattern_var,
                      value="prefix").pack(anchor="w")
    ctk.CTkRadioButton(pattern_frame, text="Add Suffix", variable=pattern_var,
                      value="suffix").pack(anchor="w")
    ctk.CTkRadioButton(pattern_frame, text="Replace Text", variable=pattern_var,
                      value="replace").pack(anchor="w")
    
    # Input fields
    input_frame = ctk.CTkFrame(dialog)
    input_frame.pack(fill="x", padx=20, pady=10)
    
    ctk.CTkLabel(input_frame, text="Text to add/replace:").pack(anchor="w")
    text_var = ctk.StringVar()
    ctk.CTkEntry(input_frame, textvariable=text_var, width=500).pack(fill="x")
    
    # Preview
    preview_frame = ctk.CTkScrollableFrame(dialog, height=150)
    preview_frame.pack(fill="both", expand=True, padx=20, pady=10)
    
    def update_preview():
        for widget in preview_frame.winfo_children():
            widget.destroy()
        
        pattern = pattern_var.get()
        text = text_var.get()
        
        for error in errors[:10]:  # Show first 10
            old_name = error.file_name
            new_name = old_name
            
            if pattern == "prefix":
                new_name = f"{text}{old_name}"
            elif pattern == "suffix":
                name, ext = os.path.splitext(old_name)
                new_name = f"{name}{text}{ext}"
            elif pattern == "replace":
                # Need find/replace fields
                pass
            
            label = ctk.CTkLabel(preview_frame, text=f"{old_name} → {new_name}",
                                anchor="w")
            label.pack(fill="x")
    
    text_var.trace_add("write", lambda *args: update_preview())
    pattern_var.trace_add("write", lambda *args: update_preview())
    
    # Buttons
    button_frame = ctk.CTkFrame(dialog)
    button_frame.pack(pady=10)
    
    ctk.CTkButton(button_frame, text="Apply Rename", width=150,
                 command=lambda: app._apply_bulk_rename(errors, pattern_var.get(), text_var.get(), dialog)).pack(side="left", padx=5)
    ctk.CTkButton(button_frame, text="Cancel", width=100,
                 command=dialog.destroy).pack(side="left", padx=5)

def apply_bulk_rename(app, errors, pattern, text, dialog):
    """Apply bulk rename operation."""
    # Implementation similar to single rename but in batch
    try:
        renamed_count = 0
        for error in errors:
            old_name = error.file_name
            new_name = old_name
            
            if pattern == "prefix":
                new_name = f"{text}{old_name}"
            elif pattern == "suffix":
                name, ext = os.path.splitext(old_name)
                new_name = f"{name}{text}{ext}"
            
            # Rename and move
            old_path = Path(error.current_location) / old_name
            new_path = Path(error.current_location) / new_name
            
            if old_path.exists():
                old_path.rename(new_path)
                renamed_count += 1
        
        # Remove from error list
        for error in errors:
            if error in app.file_errors:
                app.file_errors.remove(error)
        
        app._refresh_file_errors_display()
        dialog.destroy()
        messagebox.showinfo("Success", f"Renamed and moved {renamed_count} files.")
    
    except Exception as e:
        messagebox.showerror("Bulk Rename Failed", f"Error: {str(e)}")

def open_error_location(app):
    """Open file location for selected errors."""
    selected = [e for e in app.file_errors if e.selected]
    if not selected:
        messagebox.showwarning("No Selection", "Please select files to open location.")
        return
    
    # Open first selected location
    location = selected[0].current_location
    if Path(location).exists():
        import platform
        if platform.system() == "Windows":
            os.startfile(location)
        elif platform.system() == "Darwin":
            os.system(f'open "{location}"')
        else:
            os.system(f'xdg-open "{location}"')
    else:
        messagebox.showerror("Location Not Found", f"Could not find:\n{location}")

def mark_errors_resolved(app):
    """Mark selected errors as resolved."""
    selected = [e for e in app.file_errors if e.selected]
    if not selected:
        messagebox.showwarning("No Selection", "Please select errors to mark as resolved.")
        return
    
    for error in selected:
        if error in app.file_errors:
            app.file_errors.remove(error)
    
    app._refresh_file_errors_display()
    messagebox.showinfo("Success", f"Marked {len(selected)} errors as resolved.")

def remove_duplicates(app):
    """Remove selected duplicate/unrecognized files to Review folder."""
    selected = [e for e in app.file_errors if e.selected]
    if not selected:
        messagebox.showwarning("No Selection", "Please select files to remove.")
        return
    
    # Filter for only Duplicates and Unrecognized Format
    valid_types = ["Duplicate", "Unrecognized Format"]
    removable = [e for e in selected if e.error_type in valid_types]
    
    if not removable:
        messagebox.showwarning("Invalid Selection", 
                             "Remove button only works for 'Duplicate' and 'Unrecognized Format' errors.\n\n" +
                             "For other errors, use 'Rename Selected' to fix them.")
        return
    
    # Move files to Review folder
    base = app.base_dir.get()
    if not base:
        messagebox.showerror("Error", "Base directory not set.")
        return
    
    review_base = Path(base) / "Review_Files"
    moved_count = 0
    
    for error in removable:
        try:
            # Determine subfolder
            if error.error_type == "Duplicate":
                subfolder = "Duplicates"
            else:  # Unrecognized Format
                subfolder = "Unrecognized_Format"
            
            review_folder = review_base / subfolder
            review_folder.mkdir(parents=True, exist_ok=True)
            
            # Move file
            if error.file_path and Path(error.file_path).exists():
                dest = review_folder / error.file_name
                shutil.move(str(Path(error.file_path)), str(dest))
                moved_count += 1
                
                # Remove from error list
                app.file_errors.remove(error)
        
        except Exception as e:
            messagebox.showerror("Move Failed", f"Could not move {error.file_name}:\n{str(e)}")
    
    app._refresh_file_errors_display()
    
    if moved_count > 0:
        messagebox.showinfo("Success", 
                          f"Moved {moved_count} files to Review_Files folder.\n\n" +
                          "These files have been excluded from the import.")

def select_all_errors(app):
    """Select all errors."""
    for error in app.file_errors:
        error.selected = True
    app._refresh_file_errors_display()

def deselect_all_errors(app):
    """Deselect all errors."""
    for error in app.file_errors:
        error.selected = False
    app._refresh_file_errors_display()

def rerun_import(app):
    """Re-run import after fixing errors."""
    response = messagebox.askyesno("Re-run Import",
                                  "This will re-run the import process.\nContinue?")
    if response:
        app._execute_import()



def preview_selected_pdf(app):
    """Open a popup PDF preview of the selected file using pdf_viewer.py."""
    import importlib.util
    import sys
    import tkinter as tk
    from tkinter import ttk, messagebox as _mb
    from pathlib import Path

    # Find a selected error with a .pdf file
    selected = [e for e in app.file_errors if e.selected]
    if not selected:
        _mb.showwarning("No Selection", "Please select a file to preview.")
        return

    error = selected[0]

    # Resolve the file path
    file_path = Path(error.file_path) if error.file_path else None
    if not file_path or not file_path.exists():
        # Try building path from current_location + file_name
        if error.current_location and error.file_name:
            file_path = Path(error.current_location) / error.file_name
        if not file_path or not file_path.exists():
            _mb.showwarning("File Not Found",
                            f"Could not locate the file:\n{error.file_name}\n\n"
                            "Make sure the file still exists at its current location.")
            return

    if file_path.suffix.lower() != ".pdf":
        _mb.showinfo("Not a PDF",
                     f"{file_path.name} is not a PDF file and cannot be previewed.")
        return

    # Load pdf_viewer.py dynamically
    candidates = [
        Path(sys.argv[0]).parent / "pdf_viewer.py",
        Path.cwd() / "pdf_viewer.py",
    ]
    viewer_path = next((p for p in candidates if p.exists()), None)

    if viewer_path is None:
        _mb.showerror("Viewer Not Found",
                      "Could not find pdf_viewer.py.\n\n"
                      "Place it in the same folder as main.py to enable PDF preview.")
        return

    spec   = importlib.util.spec_from_file_location("pdf_viewer", viewer_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)

    # Open the viewer in a Toplevel (not a new Tk root)
    popup = tk.Toplevel()
    popup.title(f"PDF Preview — {file_path.name}")
    popup.geometry("900x700")
    popup.configure(bg="#1e1e2e")

    # Style scrollbars for the popup
    style = ttk.Style(popup)
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure("Vertical.TScrollbar",
                    background="#45475a", troughcolor="#1e1e2e", bordercolor="#1e1e2e")
    style.configure("Horizontal.TScrollbar",
                    background="#45475a", troughcolor="#1e1e2e", bordercolor="#1e1e2e")

    viewer = module.PDFViewer(popup)
    viewer.open_file(str(file_path))


def compare_duplicate(app, error_index):
    """Open side-by-side PDF comparison popup for a duplicate file."""
    import importlib.util
    import sys
    import tkinter as tk
    from tkinter import ttk

    error = app.file_errors[error_index]

    # ── Resolve the duplicate file path ─────────────────────────────
    dup_path = Path(error.file_path) if error.file_path else None
    if not dup_path or not dup_path.exists():
        if error.current_location and error.file_name:
            dup_path = Path(error.current_location) / error.file_name
    if not dup_path or not dup_path.exists():
        messagebox.showerror("File Not Found",
                             f"Could not locate duplicate file:\n{error.file_name}")
        return

    # ── Resolve the kept file path from issue text ───────────────────
    # Issue text looks like: "Kept: some_other_file.pdf" or the raw issue
    kept_name = ""
    raw_issue = error.issue or ""
    dry_prefix = "[DRY RUN] "
    clean = raw_issue[len(dry_prefix):] if raw_issue.startswith(dry_prefix) else raw_issue
    if "kept:" in clean.lower():
        kept_name = clean.split("kept:", 1)[-1].strip()

    kept_path = None
    if kept_name:
        # Try same folder as duplicate
        candidate = dup_path.parent / kept_name
        if candidate.exists():
            kept_path = candidate
        else:
            # Walk upward a couple levels to find it
            for parent in [dup_path.parent.parent, dup_path.parent.parent.parent]:
                try:
                    matches = list(parent.rglob(kept_name))
                    if matches:
                        kept_path = matches[0]
                        break
                except Exception:
                    pass

    if not kept_path or not kept_path.exists():
        messagebox.showwarning("Kept File Not Found",
                               f"Could not locate the kept file:\n{kept_name or '(unknown)'}\n\n"
                               "The duplicate file path will still be shown on the left.")
        kept_path = None

    # ── Load pdf_viewer.py ───────────────────────────────────────────
    candidates = [
        Path(sys.argv[0]).parent / "pdf_viewer.py",
        Path.cwd() / "pdf_viewer.py",
    ]
    viewer_path = next((p for p in candidates if p.exists()), None)

    pdf_viewer_mod = None
    if viewer_path:
        spec = importlib.util.spec_from_file_location("pdf_viewer", viewer_path)
        pdf_viewer_mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(pdf_viewer_mod)

    # ── Build popup ──────────────────────────────────────────────────
    popup = ctk.CTkToplevel(app)
    popup.title("Compare Duplicate Files")
    popup.geometry("1400x820")
    popup.configure(fg_color="#1e1e2e")
    popup.transient(app)
    popup.grab_set()

    # State tracking
    state = {"kept_side": None}   # "left" | "right"
    # Track all PDF viewer instances so we can close them before file ops
    open_viewers = []

    # Title bar
    title_bar = ctk.CTkFrame(popup, fg_color="#2b2b2b", height=44)
    title_bar.pack(fill="x")
    title_bar.pack_propagate(False)
    ctk.CTkLabel(title_bar, text="Side-by-Side Duplicate Comparison",
                 font=ctk.CTkFont(size=15, weight="bold")).pack(side="left", padx=15, pady=8)
    ctk.CTkLabel(title_bar,
                 text="Select which file to keep, then choose what to do with the other.",
                 text_color="gray", font=ctk.CTkFont(size=12)).pack(side="left", padx=5)

    # Two-column body
    body = ctk.CTkFrame(popup, fg_color="transparent")
    body.pack(fill="both", expand=True, padx=10, pady=10)
    body.columnconfigure(0, weight=1)
    body.columnconfigure(1, weight=1)
    body.rowconfigure(0, weight=1)

    def file_info_text(p: Path) -> str:
        try:
            size_kb = p.stat().st_size / 1024
            import datetime
            mtime = datetime.datetime.fromtimestamp(p.stat().st_mtime)
            return f"Size: {size_kb:.1f} KB    Modified: {mtime.strftime('%Y-%m-%d %H:%M')}"
        except Exception:
            return ""

    def build_side(parent_col, label_text, file_path: Path, side_key: str,
                   other_action_frame_ref: list):
        """Build one side panel. other_action_frame_ref is a 1-element list holding the
        other side's action frame so we can reveal it on Keep click."""

        panel = ctk.CTkFrame(parent_col, fg_color="#2b2b2b", corner_radius=8)
        panel.pack(fill="both", expand=True, padx=6, pady=4)

        # Header
        header = ctk.CTkFrame(panel, fg_color="#1e1e2e", height=36)
        header.pack(fill="x")
        header.pack_propagate(False)
        ctk.CTkLabel(header, text=label_text,
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color="#7eb8f7").pack(side="left", padx=10, pady=6)

        # File name
        fname = file_path.name if file_path else "(file not found)"
        fname_label = ctk.CTkLabel(panel, text=fname, font=ctk.CTkFont(size=12, weight="bold"),
                                   wraplength=580, anchor="w")
        fname_label.pack(fill="x", padx=10, pady=(6, 0))

        # File info
        info = file_info_text(file_path) if file_path and file_path.exists() else "File not available"
        ctk.CTkLabel(panel, text=info, font=ctk.CTkFont(size=11),
                     text_color="gray", anchor="w").pack(fill="x", padx=10, pady=(0, 4))

        # Full path
        path_str = str(file_path) if file_path else "—"
        path_label = ctk.CTkLabel(panel, text=path_str, font=ctk.CTkFont(size=10),
                                  text_color="#555577", anchor="w", wraplength=580)
        path_label.pack(fill="x", padx=10, pady=(0, 6))

        # PDF preview area
        preview_frame = ctk.CTkFrame(panel, fg_color="#111122", corner_radius=4)
        preview_frame.pack(fill="both", expand=True, padx=8, pady=4)

        if pdf_viewer_mod and file_path and file_path.exists() and file_path.suffix.lower() == ".pdf":
            try:
                inner = tk.Frame(preview_frame, bg="#111122")
                inner.pack(fill="both", expand=True)
                v = pdf_viewer_mod.PDFViewer(inner)
                v.open_file(str(file_path))
                open_viewers.append(v)  # track for cleanup before file ops
            except Exception as ex:
                ctk.CTkLabel(preview_frame, text=f"Preview unavailable:\n{ex}",
                             text_color="gray").pack(expand=True)
        elif not pdf_viewer_mod:
            ctk.CTkLabel(preview_frame, text="pdf_viewer.py not found — preview unavailable",
                         text_color="gray").pack(expand=True)
        else:
            ctk.CTkLabel(preview_frame,
                         text="File not available for preview" if not (file_path and file_path.exists())
                              else f"Not a PDF: {file_path.suffix}",
                         text_color="gray").pack(expand=True)

        # Controls container — packed at build time (fill="x", side="bottom", NOT
        # expanded) so every revealed row (keep/rename buttons, rename entry, action
        # buttons) has a reserved place below the preview. The preview frame above is
        # packed expand=True; revealing a row into `panel` after that leaves it 0px
        # (clipped) or detaches it to the root window (the stray top-left textbox).
        controls = ctk.CTkFrame(panel, fg_color="transparent")
        controls.pack(fill="x", side="bottom", padx=8, pady=6)

        # Keep button + action area
        bottom = ctk.CTkFrame(controls, fg_color="transparent")
        bottom.pack(fill="x")

        keep_btn_ref = [None]

        def on_keep():
            if state["kept_side"] is not None:
                return  # already chosen
            state["kept_side"] = side_key

            # Green highlight on this panel header
            header.configure(fg_color="#1a4a1a")
            keep_btn_ref[0].configure(text="✅ Keeping this one",
                                      fg_color="#1a4a1a", state="disabled")

            # Reveal action buttons on the OTHER side, at the bottom of that side's
            # reserved controls container (not into `panel`, which would clip them).
            other_frame = other_action_frame_ref[0]
            if other_frame:
                other_frame.pack(fill="x", side="bottom", pady=(6, 0))

        keep_btn = ctk.CTkButton(
            bottom, text="✓ Keep this one",
            fg_color="#1f538d", hover_color="#14375e",
            width=180, height=32,
            command=on_keep
        )
        keep_btn.pack(side="left", padx=4)
        keep_btn_ref[0] = keep_btn

        # ── Rename-in-place controls ─────────────────────────────────
        # Mutable holder so the current path follows a successful rename.
        cur_path = {"path": file_path}
        rename_btn_ref = [None]

        def do_rename():
            """Rename THIS file in place (stays where it was found)."""
            p = cur_path["path"]
            if not p or not p.exists():
                messagebox.showerror("File Not Found", f"Cannot find:\n{p}")
                return
            new_stem = rename_entry.get().strip()
            if not new_stem:
                messagebox.showwarning("Invalid Name", "The file name cannot be empty.")
                return
            ext = p.suffix  # preserve extension automatically
            new_path = p.parent / f"{new_stem}{ext}"
            if new_path == p:
                messagebox.showinfo("No Change", "The name is unchanged.")
                return
            if new_path.exists():
                messagebox.showwarning(
                    "Name Already Exists",
                    f"A file named:\n{new_path.name}\n\nalready exists in that folder. "
                    "Choose a different name."
                )
                return
            try:
                # Close viewer(s) for THIS file to release Windows file lock
                for viewer in open_viewers:
                    try:
                        if hasattr(viewer, 'doc') and viewer.doc:
                            viewer.doc.close()
                            viewer.doc = None
                    except Exception:
                        pass
                p.rename(new_path)
                cur_path["path"] = new_path

                # Update this side's displayed name/path + confirmation state
                fname_label.configure(text=new_path.name)
                path_label.configure(text=str(new_path))
                header.configure(fg_color="#1a4a1a")
                rename_row.pack_forget()
                rename_btn_ref[0].configure(text="✅ Renamed", state="disabled",
                                            fg_color="#1a4a1a")

                # Clear the error in the background; keep popup open so the
                # other side stays actionable.
                try:
                    app.file_errors.pop(error_index)
                    app._refresh_file_errors_display()
                except Exception:
                    pass

                messagebox.showinfo("Renamed", f"File renamed to:\n{new_path.name}")
            except Exception as ex:
                messagebox.showerror("Rename Failed", str(ex))

        def toggle_rename():
            if rename_row.winfo_ismapped():
                rename_row.pack_forget()
            else:
                rename_row.pack(fill="x", pady=(6, 0), after=bottom)
                rename_entry.delete(0, "end")
                # Pre-fill with current stem (extension excluded/preserved)
                rename_entry.insert(0, cur_path["path"].stem)
                rename_entry.focus_set()

        rename_btn = ctk.CTkButton(
            bottom, text="✏ Rename",
            fg_color="#3a3a55", hover_color="#4a4a70",
            width=120, height=32,
            command=toggle_rename
        )
        rename_btn.pack(side="left", padx=4)
        rename_btn_ref[0] = rename_btn
        if not (file_path and file_path.exists()):
            rename_btn.configure(state="disabled")

        # Rename entry row (hidden until Rename is clicked) — parent is `controls`,
        # NOT `panel`, so it reveals into reserved space instead of detaching.
        rename_row = ctk.CTkFrame(controls, fg_color="transparent")
        # Not packed yet — revealed by toggle_rename
        rename_entry = ctk.CTkEntry(rename_row, placeholder_text="New file name (no extension)")
        rename_entry.pack(side="left", fill="x", expand=True, padx=(0, 6))
        rename_entry.bind("<Return>", lambda _e: do_rename())
        ctk.CTkButton(
            rename_row, text="Confirm", width=90, height=28,
            fg_color="#1f538d", hover_color="#14375e",
            command=do_rename
        ).pack(side="left")

        # Action frame (hidden until the OTHER side is marked Keep) — parent is
        # `controls`, NOT `panel`, so its buttons aren't clipped to 0px height.
        action_frame = ctk.CTkFrame(controls, fg_color="#3a1a1a", corner_radius=6)
        # Not packed yet — revealed when other side is kept

        def do_move_other():
            """Move THIS file (not kept) to an Other folder next to its parent."""
            if not file_path or not file_path.exists():
                messagebox.showerror("File Not Found", f"Cannot find:\n{file_path}")
                return
            other_folder = file_path.parent.parent / "Other"
            dest = other_folder / file_path.name
            confirm = messagebox.askyesno(
                "Move to Other Folder",
                f"Move this file to:\n{other_folder}\n\n"
                f"File: {file_path.name}\n\n"
                "(The 'Other' folder will be created if it doesn't exist.)"
            )
            if not confirm:
                return
            try:
                other_folder.mkdir(parents=True, exist_ok=True)
                # Handle name collision
                if dest.exists():
                    base, ext = dest.stem, dest.suffix
                    i = 1
                    while dest.exists():
                        dest = other_folder / f"{base}_{i}{ext}"
                        i += 1
                # Close all PDF viewers to release Windows file locks before moving
                for viewer in open_viewers:
                    try:
                        if hasattr(viewer, 'doc') and viewer.doc:
                            viewer.doc.close()
                            viewer.doc = None
                    except Exception:
                        pass
                open_viewers.clear()
                popup.destroy()  # destroy popup (and embedded widgets) before move
                shutil.move(str(file_path), str(dest))
                app.file_errors.pop(error_index)
                app._refresh_file_errors_display()
                messagebox.showinfo("Done", f"Moved to:\n{dest}")
            except Exception as ex:
                messagebox.showerror("Move Failed", str(ex))

        def do_remove_other():
            """Move THIS file (not kept) to Review_Files/Duplicates."""
            if not file_path or not file_path.exists():
                messagebox.showerror("File Not Found", f"Cannot find:\n{file_path}")
                return
            base = app.base_dir.get()
            if not base:
                messagebox.showerror("Error", "Base directory not set.")
                return
            review_folder = Path(base) / "Review_Files" / "Duplicates"
            confirm = messagebox.askyesno(
                "Remove Duplicate",
                f"Move this file to Review_Files/Duplicates?\n\n"
                f"File: {file_path.name}"
            )
            if not confirm:
                return
            try:
                review_folder.mkdir(parents=True, exist_ok=True)
                dest = review_folder / file_path.name
                if dest.exists():
                    base_name, ext = dest.stem, dest.suffix
                    i = 1
                    while dest.exists():
                        dest = review_folder / f"{base_name}_{i}{ext}"
                        i += 1
                # Close all PDF viewers to release Windows file locks before moving
                for viewer in open_viewers:
                    try:
                        if hasattr(viewer, 'doc') and viewer.doc:
                            viewer.doc.close()
                            viewer.doc = None
                    except Exception:
                        pass
                open_viewers.clear()
                popup.destroy()  # destroy popup (and embedded widgets) before move
                shutil.move(str(file_path), str(dest))
                app.file_errors.pop(error_index)
                app._refresh_file_errors_display()
                messagebox.showinfo("Done",
                                    f"Moved to Review_Files/Duplicates:\n{file_path.name}")
            except Exception as ex:
                messagebox.showerror("Move Failed", str(ex))

        def do_move_existing():
            """Move THIS file into a folder chosen from within /meetings via a
            two-step picker: pick a top-level folder under 'meetings', then pick
            one of its immediate child folders as the destination."""
            if not file_path or not file_path.exists():
                messagebox.showerror("File Not Found", f"Cannot find:\n{file_path}")
                return

            # ── Derive the /meetings root by walking up the file's path ──────
            meetings_root = None
            for ancestor in file_path.parents:
                if ancestor.name.lower() == "meetings":
                    meetings_root = ancestor
                    break
            if meetings_root is None:
                # Fallback: the file's grandparent (…/meetings/<body>/<sub>/file)
                meetings_root = file_path.parent.parent

            try:
                top_folders = sorted(
                    [d for d in meetings_root.iterdir() if d.is_dir()],
                    key=lambda d: d.name.lower()
                )
            except Exception as ex:
                messagebox.showerror("Error", f"Could not read:\n{meetings_root}\n\n{ex}")
                return
            if not top_folders:
                messagebox.showwarning("No Folders",
                                       f"No folders found in:\n{meetings_root}")
                return

            # ── Picker dialog (single Toplevel, swapped between two steps) ───
            picker = ctk.CTkToplevel(popup)
            picker.title("Move to Existing Folder")
            picker.geometry("460x520")
            picker.configure(fg_color="#1e1e2e")
            picker.transient(popup)
            picker.grab_set()

            sel = {"top": None, "child": None}

            header_lbl = ctk.CTkLabel(picker, text="", font=ctk.CTkFont(size=13, weight="bold"),
                                      text_color="#7eb8f7", anchor="w")
            header_lbl.pack(fill="x", padx=14, pady=(12, 0))
            sub_lbl = ctk.CTkLabel(picker, text="", font=ctk.CTkFont(size=11),
                                   text_color="#888899", anchor="w")
            sub_lbl.pack(fill="x", padx=14, pady=(2, 6))

            list_frame = ctk.CTkScrollableFrame(picker, fg_color="#2b2b2b")
            list_frame.pack(fill="both", expand=True, padx=12, pady=4)

            btn_bar = ctk.CTkFrame(picker, fg_color="transparent")
            btn_bar.pack(fill="x", padx=12, pady=10)

            left_btn = ctk.CTkButton(btn_bar, text="Cancel", width=90, height=30,
                                     fg_color="#444", hover_color="#555")
            left_btn.pack(side="left")
            right_btn = ctk.CTkButton(btn_bar, text="Next", width=110, height=30,
                                      fg_color="#1f538d", hover_color="#14375e")
            right_btn.pack(side="right")

            row_refs = {"buttons": [], "selected": None}

            def render_list(folders, selected_key):
                for rb in row_refs["buttons"]:
                    rb.destroy()
                row_refs["buttons"] = []
                for d in folders:
                    is_sel = (d == selected_key)
                    b = ctk.CTkButton(
                        list_frame, text=f"  📁  {d.name}",
                        anchor="w", height=34,
                        fg_color="#85b7eb" if is_sel else "transparent",
                        text_color="#0f2438" if is_sel else "#dddddd",
                        hover_color="#3a3a55",
                        command=lambda folder=d: on_pick(folder)
                    )
                    b.pack(fill="x", padx=4, pady=2)
                    row_refs["buttons"].append(b)

            def on_pick(folder):
                if picker_state["step"] == 1:
                    sel["top"] = folder
                else:
                    sel["child"] = folder
                row_refs["selected"] = folder
                render_list(picker_state["folders"], folder)

            picker_state = {"step": 1, "folders": top_folders}

            def show_step1():
                picker_state["step"] = 1
                picker_state["folders"] = top_folders
                row_refs["selected"] = sel["top"]
                header_lbl.configure(text="Move to folder in /meetings")
                sub_lbl.configure(text="Choose a top-level folder")
                left_btn.configure(text="Cancel", command=picker.destroy)
                right_btn.configure(text="Next", command=go_next)
                render_list(top_folders, sel["top"])

            def go_next():
                if not sel["top"]:
                    messagebox.showwarning("No Selection", "Pick a folder to continue.",
                                           parent=picker)
                    return
                try:
                    children = sorted(
                        [d for d in sel["top"].iterdir() if d.is_dir()],
                        key=lambda d: d.name.lower()
                    )
                except Exception as ex:
                    messagebox.showerror("Error", f"Could not read:\n{sel['top']}\n\n{ex}",
                                         parent=picker)
                    return
                if not children:
                    messagebox.showwarning(
                        "No Child Folders",
                        f"{sel['top'].name} has no child folders to move into.",
                        parent=picker)
                    return
                picker_state["step"] = 2
                picker_state["folders"] = children
                sel["child"] = None
                row_refs["selected"] = None
                header_lbl.configure(text=sel["top"].name)
                sub_lbl.configure(text="Choose a child folder")
                left_btn.configure(text="Back", command=show_step1)
                right_btn.configure(text="Move here", command=do_final_move)
                render_list(children, None)

            def do_final_move():
                if not sel["child"]:
                    messagebox.showwarning("No Selection", "Pick a child folder.",
                                           parent=picker)
                    return
                dest_folder = sel["child"]
                dest = dest_folder / file_path.name
                confirm = messagebox.askyesno(
                    "Confirm Move",
                    f"Move this file to:\n{dest_folder}\n\nFile: {file_path.name}",
                    parent=picker)
                if not confirm:
                    return
                try:
                    if dest.exists():
                        base_name, ext = dest.stem, dest.suffix
                        i = 1
                        while dest.exists():
                            dest = dest_folder / f"{base_name}_{i}{ext}"
                            i += 1
                    # Close all PDF viewers to release Windows file locks before moving
                    for viewer in open_viewers:
                        try:
                            if hasattr(viewer, 'doc') and viewer.doc:
                                viewer.doc.close()
                                viewer.doc = None
                        except Exception:
                            pass
                    open_viewers.clear()
                    picker.destroy()
                    popup.destroy()
                    shutil.move(str(file_path), str(dest))
                    app.file_errors.pop(error_index)
                    app._refresh_file_errors_display()
                    messagebox.showinfo("Done", f"Moved to:\n{dest}")
                except Exception as ex:
                    messagebox.showerror("Move Failed", str(ex))

            show_step1()

        ctk.CTkLabel(action_frame,
                     text="What should happen to this file?",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color="#ffaaaa").pack(pady=(6, 2))

        btn_row = ctk.CTkFrame(action_frame, fg_color="transparent")
        btn_row.pack(pady=(0, 8))

        ctk.CTkButton(
            btn_row, text="📁 Move to Other Folder",
            fg_color="#555500", hover_color="#777700",
            width=180, height=30,
            command=do_move_other
        ).pack(side="left", padx=6)

        ctk.CTkButton(
            btn_row, text="📂 Move to Existing Folder",
            fg_color="#1f538d", hover_color="#14375e",
            width=190, height=30,
            command=do_move_existing
        ).pack(side="left", padx=6)

        ctk.CTkButton(
            btn_row, text="🗑 Remove Duplicate",
            fg_color="#8b0000", hover_color="#a03529",
            width=160, height=30,
            command=do_remove_other
        ).pack(side="left", padx=6)

        return action_frame

    # Build both columns, cross-linking action frames
    left_col = ctk.CTkFrame(body, fg_color="transparent")
    left_col.grid(row=0, column=0, sticky="nsew")
    right_col = ctk.CTkFrame(body, fg_color="transparent")
    right_col.grid(row=0, column=1, sticky="nsew")

    right_action_ref = [None]
    left_action_ref = [None]

    left_action = build_side(left_col,
                              "DUPLICATE FILE (flagged)",
                              dup_path, "left",
                              right_action_ref)
    left_action_ref[0] = left_action

    right_action = build_side(right_col,
                               "KEPT FILE",
                               kept_path, "right",
                               left_action_ref)
    right_action_ref[0] = right_action

    # Close button
    ctk.CTkButton(popup, text="Close", width=120,
                  fg_color="#444", hover_color="#555",
                  command=popup.destroy).pack(pady=(0, 10))
