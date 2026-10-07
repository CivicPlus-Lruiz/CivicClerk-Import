"""
Subtypes tab (Standard Import)
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
import re
import os
from collections import defaultdict
from pathlib import Path
from typing import TYPE_CHECKING, List
from gui.models import Subtype
if TYPE_CHECKING:
    from gui.meeting_import_app import MeetingImportGUI

# ================================================================
# TAB 3: SUBTYPES
# ================================================================

def build_subtypes_tab(app):
    """Build the subtypes configuration tab."""
    tab = app.tab_frames["Subtypes"]
    
    # Header
    header = ctk.CTkFrame(tab)
    header.pack(fill="x", padx=10, pady=10)
    
    ctk.CTkLabel(header, text="Subtype Aliases Configuration", 
                font=ctk.CTkFont(size=16, weight="bold")).pack(side="left", padx=5)
    
    # Action buttons
    button_frame = ctk.CTkFrame(tab)
    button_frame.pack(fill="x", padx=10, pady=5)
    
    ctk.CTkButton(button_frame, text="+ Add Subtype", width=120,
                 command=app._add_subtype).pack(side="left", padx=5)
    
    ctk.CTkButton(button_frame, text="Remove Selected", width=120,
                 command=app._remove_selected_subtypes).pack(side="left", padx=5)
    
    ctk.CTkButton(button_frame, text="Clear All", width=100,
                 fg_color="#c44536", hover_color="#a03529",
                 command=app._clear_all_subtypes).pack(side="left", padx=5)
    
    # Refresh button to update categories from Meeting Bodies
    ctk.CTkButton(button_frame, text="🔄 Refresh Categories", width=150,
                 command=app._refresh_subtypes_display,
                 fg_color="#555555", 
                 hover_color="#666666").pack(side="left", padx=5)
    
    # Helper text
    ctk.CTkLabel(button_frame, text="← Update dropdowns from Meeting Bodies (Tab 2)",
                text_color="gray", font=ctk.CTkFont(size=10)).pack(side="left", padx=5)
    
    # Advanced Alias Matching - Collapsible Section
    app._create_advanced_matching_section(tab)
    
    # Cascading note
    note_frame = ctk.CTkFrame(tab, fg_color="#2b2b2b")
    note_frame.pack(fill="x", padx=10, pady=(0, 5))
    ctk.CTkLabel(note_frame, 
                text="ℹ️  Alias Matching is cascading - When multiple subtypes match a file, the script uses the first match (top = highest priority). Use the ▲▼ arrows to set priority order.",
                text_color="#7eb8f7",
                font=ctk.CTkFont(size=11),
                anchor="w",
                wraplength=1200).pack(padx=10, pady=6, anchor="w")
    
    # Table headers with grid layout for perfect alignment
    headers_frame = ctk.CTkFrame(tab)
    headers_frame.pack(fill="x", padx=10, pady=5)
    
    # Configure columns to expand proportionally with minimum sizes
    headers_frame.grid_columnconfigure(0, minsize=55)   # Up/Down arrows
    headers_frame.grid_columnconfigure(1, minsize=30)   # Checkbox
    headers_frame.grid_columnconfigure(2, minsize=300)  # Event Category (FIXED 300px)
    headers_frame.grid_columnconfigure(3, weight=1, minsize=150)  # Folder Name
    headers_frame.grid_columnconfigure(4, weight=1, minsize=120)  # Subtype
    headers_frame.grid_columnconfigure(5, weight=2, minsize=180)  # Aliases (wider)
    headers_frame.grid_columnconfigure(6, minsize=135)  # Override Time (fixed)
    headers_frame.grid_columnconfigure(7, weight=1, minsize=180)  # Override Meeting Type
    headers_frame.grid_columnconfigure(8, weight=1, minsize=180)  # Event Category Override
    
    ctk.CTkLabel(headers_frame, text="Order", width=55, anchor="w",
                font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=2, sticky="ew")  # Order header
    ctk.CTkLabel(headers_frame, text="☐", width=30, anchor="w",
                font=ctk.CTkFont(weight="bold")).grid(row=0, column=1, padx=2, sticky="ew")
    ctk.CTkLabel(headers_frame, text="Event Category", width=300, anchor="w",
                font=ctk.CTkFont(weight="bold")).grid(row=0, column=2, padx=2, sticky="w")
    ctk.CTkLabel(headers_frame, text="Folder Name", anchor="w",
                font=ctk.CTkFont(weight="bold")).grid(row=0, column=3, padx=2, sticky="ew")
    ctk.CTkLabel(headers_frame, text="Subtype", anchor="w",
                font=ctk.CTkFont(weight="bold")).grid(row=0, column=4, padx=2, sticky="ew")
    ctk.CTkLabel(headers_frame, text="Aliases", anchor="w",
                font=ctk.CTkFont(weight="bold")).grid(row=0, column=5, padx=(14, 2), sticky="ew")
    ctk.CTkLabel(headers_frame, text="Override Time", width=135, anchor="w",
                font=ctk.CTkFont(weight="bold")).grid(row=0, column=6, padx=2, sticky="ew")
    ctk.CTkLabel(headers_frame, text="Override Meeting Type", anchor="w",
                font=ctk.CTkFont(weight="bold")).grid(row=0, column=7, padx=2, sticky="ew")
    ctk.CTkLabel(headers_frame, text="Event Category Override", anchor="w",
                font=ctk.CTkFont(weight="bold")).grid(row=0, column=8, padx=2, sticky="ew")
    
    # Scrollable subtype list
    app.subtypes_scroll = ctk.CTkScrollableFrame(tab)
    app.subtypes_scroll.pack(fill="both", expand=True, padx=10, pady=10)
    
    # Next button for Subtypes tab
    next_frame = ctk.CTkFrame(tab)
    next_frame.pack(fill="x", side="bottom", padx=10, pady=10)
    
    ctk.CTkButton(
        next_frame,
        text="Next →",
        width=120,
        command=app._next_from_subtypes,
        fg_color="#2b7a4f",
        hover_color="#1e5a3b"
    ).pack(side="right")
    
    app._refresh_subtypes_display()

def create_advanced_matching_section(app, parent):
    """Create the collapsible Advanced Alias Matching section."""
    # Container frame
    container = ctk.CTkFrame(parent, fg_color="#2b2b2b")
    container.pack(fill="x", padx=10, pady=5)
    
    # Toggle button/header
    app.advanced_matching_expanded = False
    toggle_frame = ctk.CTkFrame(container, fg_color="transparent")
    toggle_frame.pack(fill="x", padx=5, pady=5)
    
    app.advanced_toggle_label = ctk.CTkLabel(
        toggle_frame, 
        text="▶ Advanced Alias Matching",
        font=ctk.CTkFont(size=12, weight="bold"),
        cursor="hand2"
    )
    app.advanced_toggle_label.pack(side="left")
    app.advanced_toggle_label.bind("<Button-1>", lambda e: app._toggle_advanced_matching())
    
    # Content frame (initially hidden)
    app.advanced_content_frame = ctk.CTkFrame(container, fg_color="transparent")
    
    # Pattern selection row
    
    # Event Category Filter section
    filter_frame = ctk.CTkFrame(app.advanced_content_frame, fg_color="transparent")
    filter_frame.pack(fill="x", padx=10, pady=(10, 5))
    
    ctk.CTkLabel(filter_frame, text="Filter by Event Category:", 
                font=ctk.CTkFont(size=14)).pack(side="left", padx=(0, 10))
    
    # Event Category filter dropdown
    app.scan_filter_var = ctk.StringVar(value="All Meeting Bodies")
    app.scan_filter_dropdown = ctk.CTkComboBox(
        filter_frame,
        variable=app.scan_filter_var,
        values=["All Meeting Bodies"],
        width=300,
        state="readonly"
    )
    app.scan_filter_dropdown.pack(side="left", padx=(0, 10))
    
    # Refresh filter button
    ctk.CTkButton(
        filter_frame,
        text="🔄 Refresh",
        width=100,
        command=app._update_scan_filter_options,
        fg_color="#555555",
        hover_color="#666666",
        font=ctk.CTkFont(size=14)
    ).pack(side="left")
    
    pattern_frame = ctk.CTkFrame(app.advanced_content_frame, fg_color="transparent")
    pattern_frame.pack(fill="x", padx=10, pady=5)
    
    ctk.CTkLabel(pattern_frame, text="Pattern Selection:", 
                font=ctk.CTkFont(size=11)).pack(side="left", padx=(0, 10))
    
    # Dropdown for common patterns
    common_patterns = [
        "{Meeting Body}_{Subtype}_{File Type}_{Date}",
        "{Date}_{Meeting Body}_{Subtype}_{File Type}",
        "{Meeting Body}_{Subtype}_{Date}",
        "{Subtype}_{Date}_{Meeting Body}",
        "{Date} {Meeting Body} - {Subtype} {File Type}",
        "{Meeting Body}_{Meeting Body}_{Subtype}_{Subtype}_{File Type}_{Date} (Multi-word: 2-word Meeting Body, 2-word Subtype)",
        "{Meeting Body}_{Subtype}_{Subtype}_{Subtype}_{File Type}_{Date} (Multi-word: 3-word Subtype)",
        "Custom Pattern..."
    ]
    
    app.pattern_dropdown = ctk.CTkComboBox(
        pattern_frame,
        values=common_patterns,
        width=350,
        command=app._on_pattern_selected
    )
    app.pattern_dropdown.set(common_patterns[0])
    app.pattern_dropdown.pack(side="left", padx=5)
    
    ctk.CTkLabel(pattern_frame, text="or", 
                font=ctk.CTkFont(size=11)).pack(side="left", padx=10)
    
    # Custom pattern input
    app.custom_pattern_entry = ctk.CTkEntry(
        pattern_frame,
        placeholder_text="Custom Pattern: {___________________}",
        width=300
    )
    app.custom_pattern_entry.pack(side="left", padx=5)
    app.custom_pattern_entry.configure(state="disabled")  # Disabled until "Custom Pattern..." selected
    
    # Token reference note
    note_frame = ctk.CTkFrame(app.advanced_content_frame, fg_color="transparent")
    note_frame.pack(fill="x", padx=10, pady=5)
    
    ctk.CTkLabel(note_frame, 
                text="Available tokens: {Meeting Body} {Subtype} {File Type} {Date}\nRepeat tokens for multi-word items:\n  • {Meeting Body}_{Meeting Body} for \"City Council\"\n  • {Subtype}_{Subtype}_{Subtype} for \"Special Called Meeting\"",
                text_color="gray",
                font=ctk.CTkFont(size=14),
                justify="left").pack(anchor="w")
    
    # Auto-Scan button
    scan_button_frame = ctk.CTkFrame(app.advanced_content_frame, fg_color="transparent")
    scan_button_frame.pack(fill="x", padx=10, pady=10)
    
    ctk.CTkButton(
        scan_button_frame,
        text="🔍 Auto-Scan Subtypes",
        width=200,
        command=app._auto_scan_subtypes,
        fg_color="#2b7a4f",
        hover_color="#1e5a3b"
    ).pack(side="left")


def update_scan_filter_options(app):
    """Update the Event Category filter dropdown with current Meeting Bodies."""
    categories = ["All Meeting Bodies"]
    
    for body in app.meeting_bodies:
        event_category = getattr(body, 'event_category', '')
        if event_category and event_category.strip():
            if event_category not in categories:
                categories.append(event_category)
    
    # Update dropdown values
    app.scan_filter_dropdown.configure(values=categories)
    
    if len(categories) > 1:
        messagebox.showinfo("Filter Updated", f"Filter updated with {len(categories)-1} Event Categories.")
    else:
        messagebox.showinfo("No Categories", "No Event Categories found. Please configure Meeting Bodies first.")

def toggle_advanced_matching(app):
    """Toggle the Advanced Alias Matching section expanded/collapsed."""
    app.advanced_matching_expanded = not app.advanced_matching_expanded
    
    if app.advanced_matching_expanded:
        app.advanced_toggle_label.configure(text="▼ Advanced Alias Matching")
        app.advanced_content_frame.pack(fill="x", padx=5, pady=5)
    else:
        app.advanced_toggle_label.configure(text="▶ Advanced Alias Matching")
        app.advanced_content_frame.pack_forget()

def on_pattern_selected(app, choice):
    """Handle pattern dropdown selection."""
    if choice == "Custom Pattern...":
        app.custom_pattern_entry.configure(state="normal")
        app.custom_pattern_entry.delete(0, "end")
        app.custom_pattern_entry.focus()
    else:
        app.custom_pattern_entry.configure(state="disabled")
        app.custom_pattern_entry.delete(0, "end")

def auto_scan_subtypes(app):
    """Scan files and auto-populate subtypes based on selected pattern."""
    import os
    import re
    from collections import defaultdict
    
    # Get the pattern
    pattern = app.pattern_dropdown.get()
    if pattern == "Custom Pattern...":
        pattern = app.custom_pattern_entry.get().strip()
        if not pattern:
            messagebox.showwarning("No Pattern", "Please enter a custom pattern.")
            return
    
    # Get base directory from GUI
    base_dir = app.base_dir.get().strip()
    if not base_dir:
        messagebox.showwarning("No Base Directory", "Please select a base directory first.")
        return
    
    # Check for /meetings folder (case-insensitive)
    meetings_dir = None
    base_path = Path(base_dir)
    for item in base_path.iterdir():
        if item.is_dir() and item.name.lower() == "meetings":
            meetings_dir = str(item)
            break
    
    if not meetings_dir:
        messagebox.showerror("Folder Not Found", 
                           f"Could not find /meetings folder in:\n{base_dir}")
        return
    
    # Build Meeting Body lookup map (folder name -> event category)
    body_lookup = {}
    for body in app.meeting_bodies:
        # Access MeetingBody object attributes, not dict keys
        folder_path = getattr(body, 'folder_path', '') or (body.get('folder_path', '') if isinstance(body, dict) else '')
        event_category = getattr(body, 'event_category', '') or (body.get('event_category', '') if isinstance(body, dict) else '')
        
        folder_path = folder_path.strip() if folder_path else ''
        event_category = event_category.strip() if event_category else ''
        
        if folder_path and event_category:
            # Extract just the folder name (last part of path)
            folder_name = Path(folder_path).name.lower()
            body_lookup[folder_name] = event_category
    
    if not body_lookup:
        messagebox.showwarning("No Meeting Bodies", "Please configure Meeting Bodies first.")
        return
    
    # Apply Event Category filter if selected
    selected_filter = app.scan_filter_var.get()
    if selected_filter != "All Meeting Bodies":
        filtered_lookup = {
            path: cat for path, cat in body_lookup.items() 
            if cat == selected_filter
        }
        if not filtered_lookup:
            messagebox.showwarning(
                "No Matching Bodies", 
                f"No Meeting Bodies found with Event Category: {selected_filter}"
            )
            return
        body_lookup = filtered_lookup
        scan_scope = f"Event Category: {selected_filter}"
    else:
        scan_scope = "All Meeting Bodies"
    
    # Convert pattern to regex with support for repeated tokens
    import re as regex_module
    
    # Count occurrences of each token
    meeting_body_count = pattern.count("{Meeting Body}")
    subtype_count = pattern.count("{Subtype}")
    file_type_count = pattern.count("{File Type}")
    date_count = pattern.count("{Date}")
    
    # Escape the pattern first
    regex_pattern = re.escape(pattern)
    
    # Replace tokens with numbered capture groups
    for i in range(meeting_body_count):
        regex_pattern = regex_pattern.replace(
            r"\{Meeting\ Body\}", 
            f"(?P<meeting_body{i+1}>.+?)", 
            1  # Replace only first occurrence
        )
    
    for i in range(subtype_count):
        regex_pattern = regex_pattern.replace(
            r"\{Subtype\}", 
            f"(?P<subtype{i+1}>.+?)", 
            1
        )
    
    for i in range(file_type_count):
        regex_pattern = regex_pattern.replace(
            r"\{File\ Type\}", 
            f"(?P<file_type{i+1}>.+?)", 
            1
        )
    
    for i in range(date_count):
        regex_pattern = regex_pattern.replace(
            r"\{Date\}", 
            f"(?P<date{i+1}>.+?)", 
            1
        )
    
    # Replace escaped separators with flexible matching
    regex_pattern = regex_pattern.replace(r"\_", r"[_\-\s]+")
    regex_pattern = regex_pattern.replace(r"\-", r"[_\-\s]+")
    regex_pattern = regex_pattern.replace(r"\ ", r"[_\-\s]+")
    
    try:
        pattern_re = re.compile(regex_pattern, re.IGNORECASE)
    except re.error as e:
        messagebox.showerror("Invalid Pattern", f"Pattern could not be compiled: {e}")
        return

    
    # Scan files
    subtype_data = defaultdict(lambda: {"aliases": set(), "event_category": None})
    
    for root, dirs, files in os.walk(meetings_dir):
        for filename in files:
            if not filename.lower().endswith(('.pdf', '.mp4', '.mp3')):
                continue
            
            # Determine file type from folder structure
            rel_path = os.path.relpath(root, meetings_dir)
            path_parts = rel_path.split(os.sep)
            
            # Find file type folder (Agenda, Minutes, Packet, Video, etc.)
            file_type_folder = None
            for part in reversed(path_parts):
                part_lower = part.lower()
                if any(k in part_lower for k in ["agenda", "minute", "packet", "video", "notice", "caption", "other"]):
                    file_type_folder = part
                    break
            
            if not file_type_folder:
                continue
            
            # Try to match pattern
            match = pattern_re.search(filename)
            if not match:
                continue
            
            groups = match.groupdict()
            
            # Combine repeated token captures
            subtype_parts = []
            for i in range(1, 10):  # Support up to 9 repetitions
                part = groups.get(f"subtype{i}", "")
                if part:
                    subtype_parts.append(part.strip())
            
            subtype = "_".join(subtype_parts) if subtype_parts else ""
            
            if not subtype:
                continue
            
            # Look up event category from Meeting Body
            event_category = None
            # Extract the meeting body folder name (first part of relative path)
            meeting_body_folder = path_parts[0].lower() if path_parts else ""
            event_category = body_lookup.get(meeting_body_folder)
            
            if not event_category:
                continue
            
            # Store subtype and alias
            subtype_key = (subtype.lower(), event_category)
            subtype_data[subtype_key]["aliases"].add(subtype)
            subtype_data[subtype_key]["event_category"] = event_category
    
    # Add new subtypes to the list
    new_count = 0
    existing_subtypes_normalized = {
        ((s.subtype or "").strip().lower(), (s.event_category or "").strip())
        for s in app.subtypes
    }
    
    for (subtype_lower, event_category), data in subtype_data.items():
        # Skip if already exists
        if (subtype_lower, event_category) in existing_subtypes_normalized:
            continue
        
        # Add new subtype
        aliases_list = sorted(data["aliases"])
        from gui.models import Subtype
        app.subtypes.append(Subtype(
            event_category=data["event_category"],
            folder_name="",
            subtype=aliases_list[0],  # Use first alias as subtype name
            aliases=", ".join(aliases_list),
            override_time="",
            override_meeting_type="",
            event_category_override=""
        ))
        new_count += 1
    
    # Refresh display
    app._refresh_subtypes_display()
    
    # Show feedback
    messagebox.showinfo("Scan Complete", 
        f"Scanned: {scan_scope}\n\n"
        f"Found {new_count} new subtype{'s' if new_count != 1 else ''}.")

def next_from_meeting_bodies(app):
    """Navigate from Meeting Bodies to Subtypes with validation."""
    # Check if any meeting body is missing default time
    missing_time = []
    for body in app.meeting_bodies:
        if not body.default_time:
            event_cat = body.event_category if hasattr(body, 'event_category') else "Unknown"
            missing_time.append(event_cat)
    
    if missing_time:
        # Show warning
        response = messagebox.askquestion(
            "Missing Default Time",
            f"⚠️ Some Meeting Bodies are missing a default time:\n\n" +
            "\n".join(f"  • {name}" for name in missing_time[:5]) +
            ("\n  ..." if len(missing_time) > 5 else "") +
            f"\n\nThis may cause issues during import.\n\nContinue anyway?",
            icon='warning'
        )
        if response == 'no':
            return
    
    # Switch to Subtypes tab
    app._switch_tab("Subtypes")

def next_from_subtypes(app):
    """Navigate from Subtypes to Run."""
    app._switch_tab("Run")

def add_subtype(app):
    """Add a new subtype row."""
    app.subtypes.append(Subtype())
    app._refresh_subtypes_display()
    app._save_config()

def remove_selected_subtypes(app):
    """Remove selected subtypes."""
    to_remove = [i for i, st in enumerate(app.subtypes) if hasattr(st, '_selected') and st._selected]
    for idx in reversed(to_remove):
        app.subtypes.pop(idx)
    app._refresh_subtypes_display()
    app._save_config()

def clear_all_subtypes(app):
    """Clear all subtypes after confirmation."""
    if not app.subtypes:
        messagebox.showinfo("Nothing to Clear", "No subtypes to clear.")
        return
    
    response = messagebox.askyesno(
        "Clear All Subtypes?",
        f"This will remove all {len(app.subtypes)} subtypes.\n\nAre you sure?"
    )
    if response:
        app.subtypes.clear()
        app._refresh_subtypes_display()
        app._save_config()
        messagebox.showinfo("Cleared", "All subtypes have been cleared.")

def refresh_subtypes_display(app):
    """Rebuild the subtypes list display."""
    # Clear existing
    for widget in app.subtypes_scroll.winfo_children():
        widget.destroy()
    
    # Create row for each subtype
    for idx, subtype in enumerate(app.subtypes):
        # Alternating row colors
        row_color = "#2b2b2b" if idx % 2 == 0 else "#333333"
        
        row = ctk.CTkFrame(app.subtypes_scroll, fg_color=row_color)
        row.pack(fill="x", pady=2)
        row._original_color = row_color
        row._row_index = idx
        
        # Configure columns - now with drag handle column 0
        row.grid_columnconfigure(0, minsize=55)   # Up/Down arrows
        row.grid_columnconfigure(1, minsize=30)   # Checkbox
        row.grid_columnconfigure(2, minsize=300)  # Event Category (FIXED 300px)
        row.grid_columnconfigure(3, weight=1, minsize=150)  # Folder Name
        row.grid_columnconfigure(4, weight=1, minsize=120)  # Subtype
        row.grid_columnconfigure(5, weight=2, minsize=180)  # Aliases
        row.grid_columnconfigure(6, minsize=135)  # Override Time (fixed)
        row.grid_columnconfigure(7, weight=1, minsize=180)  # Override Meeting Type
        row.grid_columnconfigure(8, weight=1, minsize=180)  # Event Category Override
        
        # Up/Down arrow buttons for reordering
        arrows_frame = ctk.CTkFrame(row, fg_color="transparent")
        arrows_frame.grid(row=0, column=0, padx=2, sticky="w")
        
        up_btn = ctk.CTkButton(arrows_frame, text="▲", width=22, height=18,
                               font=ctk.CTkFont(size=10),
                               fg_color="transparent", hover_color="#444444",
                               command=lambda i=idx: app._move_subtype(i, -1))
        up_btn.pack(side="left", padx=1)
        
        down_btn = ctk.CTkButton(arrows_frame, text="▼", width=22, height=18,
                                 font=ctk.CTkFont(size=10),
                                 fg_color="transparent", hover_color="#444444",
                                 command=lambda i=idx: app._move_subtype(i, 1))
        down_btn.pack(side="left", padx=1)
        
        # Disable up on first row, down on last row
        if idx == 0:
            up_btn.configure(state="disabled", text_color="gray")
        if idx == len(app.subtypes) - 1:
            down_btn.configure(state="disabled", text_color="gray")
        
        # Checkbox
        var = ctk.BooleanVar(value=False)
        checkbox = ctk.CTkCheckBox(row, text="", variable=var, width=30,
                                   command=lambda i=idx, v=var: setattr(app.subtypes[i], '_selected', v.get()))
        checkbox.grid(row=0, column=1, padx=2, sticky="w")
        
        # Event Category (Dropdown from Meeting Bodies) - fixed width
        ec_var = ctk.StringVar(value=subtype.event_category)
        available_categories = app._get_event_categories()
        if subtype.event_category and subtype.event_category not in available_categories:
            available_categories.insert(1, subtype.event_category)
        
        ec_dropdown = ctk.CTkOptionMenu(
            row, 
            variable=ec_var, 
            values=available_categories,
            width=300,
            dynamic_resizing=False,
            command=lambda val, i=idx: (setattr(app.subtypes[i], 'event_category', val), app._save_config())
        )
        ec_dropdown.grid(row=0, column=2, padx=2, sticky="w")
        
        # Folder Name
        fn_var = ctk.StringVar(value=subtype.folder_name)
        fn_entry = ctk.CTkEntry(row, textvariable=fn_var)
        fn_entry.grid(row=0, column=3, padx=2, sticky="ew")
        fn_var.trace_add("write", lambda *args, i=idx, v=fn_var:
                       (setattr(app.subtypes[i], 'folder_name', v.get()), app._save_config())[0])
        
        # Subtype
        st_var = ctk.StringVar(value=subtype.subtype)
        st_entry = ctk.CTkEntry(row, textvariable=st_var)
        st_entry.grid(row=0, column=4, padx=2, sticky="ew")
        st_var.trace_add("write", lambda *args, i=idx, v=st_var:
                       (setattr(app.subtypes[i], 'subtype', v.get()), app._save_config())[0])
        
        # Aliases
        al_var = ctk.StringVar(value=subtype.aliases)
        al_entry = ctk.CTkEntry(row, textvariable=al_var)
        al_entry.grid(row=0, column=5, padx=2, sticky="ew")
        al_var.trace_add("write", lambda *args, i=idx, v=al_var:
                       (setattr(app.subtypes[i], 'aliases', v.get()), app._save_config())[0])
        
        # Override Time - Time Picker (fixed width)
        time_frame = ctk.CTkFrame(row, fg_color="transparent")
        time_frame.grid(row=0, column=6, padx=2, sticky="w")
        
        # Parse existing time or use empty
        if subtype.override_time and ":" in subtype.override_time:
            hour = subtype.override_time.split(":")[0]
            minute = subtype.override_time.split(":")[1]
        else:
            hour = ""
            minute = ""
        
        # Hour dropdown (empty or 00-23)
        hour_var = ctk.StringVar(value=hour)
        hour_dropdown = ctk.CTkOptionMenu(
            time_frame,
            variable=hour_var,
            values=[""] + [f"{h:02d}" for h in range(24)],
            width=60,
            command=lambda val, i=idx: app._update_subtype_time(i)
        )
        hour_dropdown.pack(side="left", padx=1)
        
        ctk.CTkLabel(time_frame, text=":", font=ctk.CTkFont(size=14, weight="bold")).pack(side="left")
        
        # Minute dropdown (empty or 00, 15, 30, 45)
        minute_var = ctk.StringVar(value=minute)
        minute_dropdown = ctk.CTkOptionMenu(
            time_frame,
            variable=minute_var,
            values=["", "00", "15", "30", "45"],
            width=60,
            command=lambda val, i=idx: app._update_subtype_time(i)
        )
        minute_dropdown.pack(side="left", padx=1)
        
        # Store references for updating
        subtype._hour_var = hour_var
        subtype._minute_var = minute_var
        
        # Override Meeting Type
        omt_var = ctk.StringVar(value=subtype.override_meeting_type)
        omt_entry = ctk.CTkEntry(row, textvariable=omt_var)
        omt_entry.grid(row=0, column=7, padx=2, sticky="ew")
        omt_var.trace_add("write", lambda *args, i=idx, v=omt_var:
                         (setattr(app.subtypes[i], 'override_meeting_type', v.get()), app._save_config())[0])
        
        # Event Category Override
        eco_var = ctk.StringVar(value=subtype.event_category_override)
        eco_entry = ctk.CTkEntry(row, textvariable=eco_var)
        eco_entry.grid(row=0, column=8, padx=2, sticky="ew")
        eco_var.trace_add("write", lambda *args, i=idx, v=eco_var:
                         setattr(app.subtypes[i], 'event_category_override', v.get()))

def update_subtype_time(app, index: int):
    """Update subtype override time from hour/minute dropdowns."""
    if 0 <= index < len(app.subtypes):
        subtype = app.subtypes[index]
        if hasattr(subtype, '_hour_var') and hasattr(subtype, '_minute_var'):
            hour = subtype._hour_var.get()
            minute = subtype._minute_var.get()
            # Only set if both are selected
            if hour and minute:
                subtype.override_time = f"{hour}:{minute}"
            else:
                subtype.override_time = ""
            app._save_config()

def move_subtype(app, index: int, direction: int):
    """Move a subtype row up (-1) or down (+1)."""
    new_index = index + direction
    if 0 <= new_index < len(app.subtypes):
        app.subtypes[index], app.subtypes[new_index] = app.subtypes[new_index], app.subtypes[index]
        app._refresh_subtypes_display()

def drag_release(app, event, index):
    """Complete the drag - reorder subtypes."""
    target_index = app._drag_data.get("target")
    source_index = app._drag_data.get("index")

    # Reset all row colors before refresh
    for widget in app.subtypes_scroll.winfo_children():
        if hasattr(widget, '_row_index'):
            widget.configure(fg_color=widget._original_color)

    # Reset drag state
    app._drag_data = {"index": None, "widget": None, "target": None, "last_target": None}

    # Reorder if valid target
    if target_index is not None and target_index != source_index:
        subtype = app.subtypes.pop(source_index)
        app.subtypes.insert(target_index, subtype)

    # Refresh display only once on release
    app._refresh_subtypes_display()

