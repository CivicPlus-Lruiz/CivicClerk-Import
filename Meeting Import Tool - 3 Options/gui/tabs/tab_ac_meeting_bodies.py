"""
Agenda Center - Meeting Bodies tab
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
from pathlib import Path
from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from gui.meeting_import_app import MeetingImportGUI

# ================================================================
# AGENDA CENTER: TAB - MEETING BODIES
# ================================================================

def build_ac_meeting_bodies_tab(app):
    """Build the Agenda Center Meeting Bodies tab."""
    tab = app.tab_frames["AC Meeting Bodies"]

    # ── Header bar with Run button ────────────────────────────────
    header = ctk.CTkFrame(tab, fg_color="#2b2b2b", height=56)
    header.pack(fill="x")
    header.pack_propagate(False)

    ctk.CTkLabel(
        header, text="Step 2: Review Event Categories & Meeting Types",
        font=ctk.CTkFont(size=15, weight="bold"),
    ).pack(side="left", padx=16, pady=12)

    ctk.CTkButton(
        header, text="▶  Run", width=160, height=38,
        fg_color="#1f538d", hover_color="#14375e",
        font=ctk.CTkFont(size=14, weight="bold"),
        command=app._ac_run_import,
    ).pack(side="right", padx=16, pady=8)

    # ── Scrollable content ────────────────────────────────────────
    app.ac_mb_scroll = ctk.CTkScrollableFrame(tab)
    app.ac_mb_scroll.pack(fill="both", expand=True, padx=12, pady=8)

    # Placeholder until file is loaded
    app.ac_mb_placeholder = ctk.CTkLabel(
        app.ac_mb_scroll,
        text="Upload a file on the Event Name tab first.",
        text_color="#888888", font=ctk.CTkFont(size=14),
    )
    app.ac_mb_placeholder.pack(pady=40)

def ac_preload_meeting_bodies(app):
    """Populate AC Meeting Bodies tab from the loaded DataFrame."""
    if app.ac_df is None:
        return

    # Clear existing content
    for widget in app.ac_mb_scroll.winfo_children():
        widget.destroy()

    app.ac_time_hour_vars = {}
    app.ac_time_min_vars = {}
    app.ac_category_rename_vars = {}
    app.ac_type_rename_vars = {}

    df = app.ac_df
    unique_categories = sorted(df["Event Category"].dropna().unique().tolist())
    unique_types = sorted(df["Meeting Type"].dropna().unique().tolist())

    # ── Section: Event Categories ─────────────────────────────────
    cat_section = ctk.CTkFrame(app.ac_mb_scroll, fg_color="#1a2535", corner_radius=10)
    cat_section.pack(fill="x", pady=(0, 12), padx=2)

    ctk.CTkLabel(
        cat_section,
        text="Event Categories",
        font=ctk.CTkFont(size=16, weight="bold"),
        text_color="#5b9bd5",
    ).pack(anchor="w", padx=16, pady=(12, 4))

    ctk.CTkLabel(
        cat_section,
        text="Rename categories as needed. Set a default time for each (used when source time is 00:00:00).",
        text_color="#888888", font=ctk.CTkFont(size=12),
    ).pack(anchor="w", padx=16, pady=(0, 8))

    # Single shared grid container — header + all rows share the same column widths
    cat_grid = ctk.CTkFrame(cat_section, fg_color="transparent")
    cat_grid.pack(fill="x", padx=12, pady=(0, 8))
    cat_grid.grid_columnconfigure(0, weight=2, minsize=200)   # Original Name
    cat_grid.grid_columnconfigure(1, weight=2, minsize=200)   # Rename To
    cat_grid.grid_columnconfigure(2, minsize=170)             # Default Time (fixed)

    # Header row (row 0)
    hdr_bg = ctk.CTkFrame(cat_grid, fg_color="#243040", height=28, corner_radius=4)
    hdr_bg.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 4))
    hdr_bg.grid_columnconfigure(0, weight=2, minsize=200)
    hdr_bg.grid_columnconfigure(1, weight=2, minsize=200)
    hdr_bg.grid_columnconfigure(2, minsize=170)
    hdr_bg.grid_propagate(False)

    ctk.CTkLabel(hdr_bg, text="Original Name", anchor="w",
                 font=ctk.CTkFont(size=11, weight="bold"),
                 text_color="#aaaaaa").grid(row=0, column=0, padx=10, sticky="w", pady=4)
    ctk.CTkLabel(hdr_bg, text="Rename To", anchor="w",
                 font=ctk.CTkFont(size=11, weight="bold"),
                 text_color="#aaaaaa").grid(row=0, column=1, padx=10, sticky="w", pady=4)
    ctk.CTkLabel(hdr_bg, text="Default Time", anchor="w",
                 font=ctk.CTkFont(size=11, weight="bold"),
                 text_color="#aaaaaa").grid(row=0, column=2, padx=10, sticky="w", pady=4)

    # Data rows — each cell placed directly into cat_grid so columns are shared
    for row_idx, cat in enumerate(unique_categories, start=1):
        bg = "#1e2a3a" if row_idx % 2 == 0 else "transparent"

        # Original label
        ctk.CTkLabel(
            cat_grid, text=cat, anchor="w",
            font=ctk.CTkFont(size=12), text_color="#dddddd",
            fg_color=bg,
        ).grid(row=row_idx, column=0, padx=10, pady=3, sticky="ew")

        # Rename entry
        rename_var = ctk.StringVar(value=app.ac_category_renames.get(cat, cat))
        app.ac_category_rename_vars[cat] = rename_var
        ctk.CTkEntry(cat_grid, textvariable=rename_var, height=30).grid(
            row=row_idx, column=1, padx=10, pady=3, sticky="ew"
        )

        # Time picker (hour : minute)
        time_frame = ctk.CTkFrame(cat_grid, fg_color=bg)
        time_frame.grid(row=row_idx, column=2, padx=10, pady=3, sticky="w")

        saved_time = app.ac_default_times.get(cat, "")
        saved_hour = saved_time.split(":")[0] if saved_time else "06"
        saved_min  = saved_time.split(":")[1] if saved_time and ":" in saved_time else "00"

        hour_var = ctk.StringVar(value=saved_hour)
        min_var  = ctk.StringVar(value=saved_min)
        app.ac_time_hour_vars[cat] = hour_var
        app.ac_time_min_vars[cat]  = min_var

        hours   = [f"{h:02d}" for h in range(0, 24)]
        minutes = ["00", "15", "30", "45"]

        ctk.CTkOptionMenu(
            time_frame, variable=hour_var, values=hours, width=72,
            command=lambda v, c=cat: app._ac_save_time(c),
        ).pack(side="left")
        ctk.CTkLabel(time_frame, text=":", font=ctk.CTkFont(size=14, weight="bold")).pack(side="left", padx=2)
        ctk.CTkOptionMenu(
            time_frame, variable=min_var, values=minutes, width=72,
            command=lambda v, c=cat: app._ac_save_time(c),
        ).pack(side="left")

    # ── Section: Meeting Types ────────────────────────────────────
    type_section = ctk.CTkFrame(app.ac_mb_scroll, fg_color="#1a2535", corner_radius=10)
    type_section.pack(fill="x", pady=(0, 12), padx=2)

    ctk.CTkLabel(
        type_section,
        text="Meeting Types",
        font=ctk.CTkFont(size=16, weight="bold"),
        text_color="#5b9bd5",
    ).pack(anchor="w", padx=16, pady=(12, 4))

    ctk.CTkLabel(
        type_section,
        text="Rename meeting types to match what is configured in CivicClerk.",
        text_color="#888888", font=ctk.CTkFont(size=12),
    ).pack(anchor="w", padx=16, pady=(0, 8))

    # Single shared grid for Meeting Types — header + rows share column widths
    type_grid = ctk.CTkFrame(type_section, fg_color="transparent")
    type_grid.pack(fill="x", padx=12, pady=(0, 8))
    type_grid.grid_columnconfigure(0, weight=1, minsize=200)
    type_grid.grid_columnconfigure(1, weight=1, minsize=200)

    # Header row
    type_hdr_bg = ctk.CTkFrame(type_grid, fg_color="#243040", height=28, corner_radius=4)
    type_hdr_bg.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 4))
    type_hdr_bg.grid_columnconfigure(0, weight=1, minsize=200)
    type_hdr_bg.grid_columnconfigure(1, weight=1, minsize=200)
    type_hdr_bg.grid_propagate(False)

    ctk.CTkLabel(type_hdr_bg, text="Original Name", anchor="w",
                 font=ctk.CTkFont(size=11, weight="bold"),
                 text_color="#aaaaaa").grid(row=0, column=0, padx=10, sticky="w", pady=4)
    ctk.CTkLabel(type_hdr_bg, text="Rename To", anchor="w",
                 font=ctk.CTkFont(size=11, weight="bold"),
                 text_color="#aaaaaa").grid(row=0, column=1, padx=10, sticky="w", pady=4)

    # Data rows placed directly into type_grid
    for row_idx, mtype in enumerate(unique_types, start=1):
        bg = "#1e2a3a" if row_idx % 2 == 0 else "transparent"

        ctk.CTkLabel(
            type_grid, text=mtype, anchor="w",
            font=ctk.CTkFont(size=12), text_color="#dddddd",
            fg_color=bg,
        ).grid(row=row_idx, column=0, padx=10, pady=3, sticky="ew")

        rename_var = ctk.StringVar(value=app.ac_type_renames.get(mtype, mtype))
        app.ac_type_rename_vars[mtype] = rename_var
        ctk.CTkEntry(type_grid, textvariable=rename_var, height=30).grid(
            row=row_idx, column=1, padx=10, pady=3, sticky="ew"
        )

def ac_save_time(app, category: str):
    """Save the default time for a category from the dropdowns."""
    hour = app.ac_time_hour_vars.get(category, ctk.StringVar(value="06")).get()
    mins = app.ac_time_min_vars.get(category, ctk.StringVar(value="00")).get()
    app.ac_default_times[category] = f"{hour}:{mins}"

def ac_run_import(app):
    """Execute the Agenda Center import: clean data and write Meetings_Output.xlsx."""
    if app.ac_df is None:
        messagebox.showwarning("No File", "Please upload an Excel file on the Event Name tab first.")
        return

    import pandas as pd

    try:
        df = app.ac_df.copy()

        # ── 1. Collect renames from UI vars ──────────────────────
        app.ac_category_renames = {
            orig: var.get()
            for orig, var in app.ac_category_rename_vars.items()
        }
        app.ac_type_renames = {
            orig: var.get()
            for orig, var in app.ac_type_rename_vars.items()
        }
        # Save times
        for cat in list(app.ac_time_hour_vars.keys()):
            app._ac_save_time(cat)

        # ── 2. Apply Event Name edits ─────────────────────────────
        for idx, new_name in app.ac_event_name_edits.items():
            df.at[idx, "Event Name"] = new_name

        # ── 3. Fix Event Date → MM/DD/YYYY string ─────────────────
        def fix_date(val):
            if pd.isna(val) or val == "":
                return ""
            try:
                return pd.to_datetime(str(val)).strftime("%m/%d/%Y")
            except Exception:
                return str(val)

        df["Event Date"] = df["Event Date"].apply(fix_date)

        # ── 4. Fix Event Time → HH:MM:SS AM/PM ───────────────────
        def fix_time(row):
            raw = str(row.get("Event Time", "") or "").strip()
            cat = str(row.get("Event Category", "") or "").strip()

            is_zero = raw in ("00:00:00", "0:00:00", "", "nan")
            if is_zero:
                # Use user-set default time for this category
                default = app.ac_default_times.get(cat, "")
                if not default:
                    return "12:00:00 AM"
                parts = default.split(":")
                h = int(parts[0]) if parts[0].isdigit() else 0
                m = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 0
                # Convert to 12-hour
                period = "AM" if h < 12 else "PM"
                h12 = h % 12 or 12
                return f"{h12:02d}:{m:02d}:00 {period}"
            else:
                # Parse 24-hour HH:MM:SS
                try:
                    parts = raw.split(":")
                    h = int(parts[0])
                    m = int(parts[1]) if len(parts) > 1 else 0
                    s = int(parts[2]) if len(parts) > 2 else 0
                    period = "AM" if h < 12 else "PM"
                    h12 = h % 12 or 12
                    return f"{h12:02d}:{m:02d}:{s:02d} {period}"
                except Exception:
                    return raw

        df["Event Time"] = df.apply(fix_time, axis=1)

        # ── 5. Rename Event Categories ────────────────────────────
        df["Event Category"] = df["Event Category"].apply(
            lambda v: app.ac_category_renames.get(str(v), str(v)) if pd.notna(v) else v
        )

        # ── 6. Rename Meeting Types ───────────────────────────────
        df["Meeting Type"] = df["Meeting Type"].apply(
            lambda v: app.ac_type_renames.get(str(v), str(v)) if pd.notna(v) else v
        )

        # ── 7. Move AdditionalVideoLink → External Media URL ──────
        skipped_count = 0
        if "AdditionalVideoLink" in df.columns and "External Media URL" in df.columns:
            for idx, row in df.iterrows():
                ext = str(row.get("External Media URL", "") or "").strip()
                add = str(row.get("AdditionalVideoLink", "") or "").strip()
                if add and not ext:
                    df.at[idx, "External Media URL"] = add
                elif add and ext:
                    skipped_count += 1

        # ── 8. Save output ────────────────────────────────────────
        # Determine output folder (same folder as source file)
        source_dir = Path(app.ac_source_file).parent
        out_path = source_dir / "Meetings_Output.xlsx"

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile="Meetings_Output.xlsx",
            initialdir=str(source_dir),
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Save Output File",
        )
        if not save_path:
            return

        df.to_excel(save_path, index=False)

        # ── 9. Popup if any rows were skipped ─────────────────────
        if skipped_count > 0:
            messagebox.showinfo(
                "AdditionalVideoLink — Rows Skipped",
                f"{skipped_count} row(s) were NOT updated because both\n"
                f"'External Media URL' and 'AdditionalVideoLink' had data.\n\n"
                f"The existing External Media URL values were kept.\n"
                f"Please review those rows manually if needed."
            )

        messagebox.showinfo(
            "Import Complete",
            f"✓ Meetings_Output.xlsx saved successfully!\n\n"
            f"Location: {save_path}"
        )

    except Exception as e:
        messagebox.showerror("Run Error", f"An error occurred during import:\n\n{str(e)}")

# ================================================================
# CONFIGURATION PERSISTENCE
# ================================================================

def save_config(app):
    """Save configuration to file in the base directory."""
    base_dir = app.base_dir.get().strip()
    if not base_dir:
        return  # Can't save without a base directory
    
    # Save config in the customer's base directory
    config_path = Path(base_dir) / "meeting_import_config.json"
    
    config = {
        "base_dir": base_dir,
        "date_format": app.date_format.get(),
        "meeting_bodies": [body.to_dict() for body in app.meeting_bodies],
        "subtypes": [st.to_dict() for st in app.subtypes],
        "checklist_state": app.checklist_state
    }
    
    try:
        with open(config_path, 'w') as f:
            json.dump(config, f, indent=2)
    except Exception as e:
        print(f"Error saving config: {e}")

def load_config(app):
    """Load configuration from file in the base directory."""
    base_dir = app.base_dir.get().strip()
    if not base_dir:
        # Try to load from global config to get the last base_dir
        global_config_path = Path.home() / ".meeting_import_gui_last_basedir.json"
        if global_config_path.exists():
            try:
                with open(global_config_path, 'r') as f:
                    global_config = json.load(f)
                last_base_dir = global_config.get("last_base_dir", "")
                if last_base_dir:
                    app.base_dir.set(last_base_dir)
                    base_dir = last_base_dir
            except Exception:
                pass
    
    if not base_dir:
        return
    
    # Load config from the customer's base directory
    config_path = Path(base_dir) / "meeting_import_config.json"
    
    if not config_path.exists():
        return
    
    try:
        with open(config_path, 'r') as f:
            config = json.load(f)
        
        app.base_dir.set(config.get("base_dir", base_dir))
        app.date_format.set(config.get("date_format", "US DATE"))
        
        app.meeting_bodies = [MeetingBody.from_dict(d) for d in config.get("meeting_bodies", [])]
        app.subtypes = [Subtype.from_dict(d) for d in config.get("subtypes", [])]
        app.checklist_state = config.get("checklist_state", {})
        
        app._refresh_meeting_bodies_display()
        app._refresh_subtypes_display()
        
        # Save this as the last used base_dir
        global_config_path = Path.home() / ".meeting_import_gui_last_basedir.json"
        try:
            with open(global_config_path, 'w') as f:
                json.dump({"last_base_dir": base_dir}, f)
        except Exception:
            pass
    
    except Exception as e:
        print(f"Error loading config: {e}")
