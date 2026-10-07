"""
Checklist tab
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
import webbrowser
from pathlib import Path
from typing import TYPE_CHECKING
from gui.constants import CHECKLIST_ITEMS
if TYPE_CHECKING:
    from gui.meeting_import_app import MeetingImportGUI

# ================================================================
# TAB 1: CHECKLIST
# ================================================================

def build_checklist_tab(app):
    """Build the checklist tab."""
    tab = app.tab_frames["Checklist"]
    
    # Site ID input at top
    header = ctk.CTkFrame(tab)
    header.pack(fill="x", padx=10, pady=(10, 5))
    
    ctk.CTkLabel(header, text="Customer Site ID:", 
                font=ctk.CTkFont(size=14, weight="bold")).pack(side="left", padx=5)
    
    site_entry = ctk.CTkEntry(header, textvariable=app.site_id, width=200,
                             placeholder_text="e.g., seattle, portland")
    site_entry.pack(side="left", padx=5)
    
    # Bind to load checklist state
    app.site_id.trace_add("write", app._on_siteid_changed)
    
    # Scrollable checklist
    app.checklist_scroll = ctk.CTkScrollableFrame(tab)
    app.checklist_scroll.pack(fill="both", expand=True, padx=10, pady=10)
    
    # Build checklist items
    app._build_checklist_items()

def build_checklist_items(app):
    """Build the checklist items with collapsible sections."""
    app.checklist_widgets = []
    item_counter = 0
    
    for section_data in CHECKLIST_ITEMS:
        # Section header (collapsible)
        section_frame = ctk.CTkFrame(app.checklist_scroll)
        section_frame.pack(fill="x", pady=(10, 5))
        
        section_label = ctk.CTkLabel(section_frame, text=section_data["section"],
                                    font=ctk.CTkFont(size=16, weight="bold"),
                                    anchor="w")
        section_label.pack(side="left", padx=10, pady=5)
        
        # Items container
        items_frame = ctk.CTkFrame(app.checklist_scroll, fg_color="transparent")
        items_frame.pack(fill="x", padx=20, pady=5)
        
        for item_data in section_data["items"]:
            item_counter += 1
            
            # Skip instruction items (no checkbox)
            if item_data.get("type") in ["instruction", "celebration"]:
                label = ctk.CTkLabel(items_frame, text=f"  {item_data['text']}",
                                    anchor="w", wraplength=800)
                label.pack(fill="x", pady=2)
                continue
            
            # Checklist item with checkbox
            item_frame = ctk.CTkFrame(items_frame, fg_color="transparent")
            item_frame.pack(fill="x", pady=2)
            
            # Checkbox
            var = ctk.BooleanVar(value=False)
            checkbox = ctk.CTkCheckBox(item_frame, text="", variable=var, width=20,
                                      command=lambda idx=item_counter: app._save_checklist_state())
            checkbox.pack(side="left", padx=(0, 5))
            
            # Text (with link if applicable)
            if item_data.get("url"):
                # Clickable link
                link_label = ctk.CTkLabel(item_frame, text=f"🔗 {item_data['text']}",
                                         anchor="w", text_color="#00BFFF",
                                         cursor="hand2",
                                         font=ctk.CTkFont(size=12, underline=True))
                link_label.pack(side="left", fill="x", expand=True)
                
                url = item_data["url"]
                needs_siteid = item_data.get("needs_siteid", False)
                link_label.bind("<Button-1>", 
                               lambda e, u=url, ns=needs_siteid: app._open_link(u, ns))
            else:
                # Regular text
                label = ctk.CTkLabel(item_frame, text=item_data["text"], anchor="w")
                label.pack(side="left", fill="x", expand=True)
            
            # Store reference
            app.checklist_widgets.append((item_counter, var))

def open_link(app, url: str, needs_siteid: bool):
    """Open a hyperlink, substituting SiteID if needed."""
    if needs_siteid:
        siteid = app.site_id.get().strip()
        if not siteid:
            messagebox.showerror("Site ID Required", 
                               "Please enter a Site ID before opening this link.")
            return
        url = url.replace("{SITEID}", siteid.lower())
    
    webbrowser.open(url)

def on_siteid_changed(app, *args):
    """Load checklist state when Site ID changes."""
    siteid = app.site_id.get().strip().lower()
    if siteid and siteid in app.checklist_state:
        state = app.checklist_state[siteid]
        for item_idx, var in app.checklist_widgets:
            var.set(state.get(item_idx, False))

def save_checklist_state(app):
    """Save current checklist state for this Site ID."""
    siteid = app.site_id.get().strip().lower()
    if not siteid:
        return
    
    state = {}
    for item_idx, var in app.checklist_widgets:
        state[item_idx] = var.get()
    
    app.checklist_state[siteid] = state
    
    # Persist to file
    app._save_config()

