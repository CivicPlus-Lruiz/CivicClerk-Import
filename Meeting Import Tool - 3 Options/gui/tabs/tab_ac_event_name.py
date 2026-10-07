"""
Agenda Center - Event Name tab
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
from pathlib import Path
from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from gui.meeting_import_app import MeetingImportGUI

# ================================================================
# AGENDA CENTER: TAB - EVENT NAME
# ================================================================

def build_ac_event_name_tab(app):
    """Build the Agenda Center Event Name tab with file upload and 99+ char review."""
    tab = app.tab_frames["AC Event Name"]

    # ── Header bar ───────────────────────────────────────────────
    header = ctk.CTkFrame(tab, fg_color="#2b2b2b", height=56)
    header.pack(fill="x", padx=0, pady=0)
    header.pack_propagate(False)

    ctk.CTkLabel(
        header, text="Step 1: Upload Excel File & Review Long Event Names",
        font=ctk.CTkFont(size=15, weight="bold"),
    ).pack(side="left", padx=16, pady=12)

    # ── Upload section ────────────────────────────────────────────
    upload_frame = ctk.CTkFrame(tab)
    upload_frame.pack(fill="x", padx=12, pady=(10, 6))

    ctk.CTkLabel(upload_frame, text="Agenda Center Excel File:",
                 font=ctk.CTkFont(size=13, weight="bold")).pack(side="left", padx=8)

    app.ac_file_label = ctk.CTkLabel(
        upload_frame, text="No file selected",
        text_color="#888888", font=ctk.CTkFont(size=13),
    )
    app.ac_file_label.pack(side="left", padx=8)

    ctk.CTkButton(
        upload_frame, text="📂 Browse...", width=130, height=34,
        command=app._ac_browse_file,
    ).pack(side="right", padx=8)

    # ── Info bar ─────────────────────────────────────────────────
    app.ac_info_bar = ctk.CTkFrame(tab, fg_color="#1a3a1a", corner_radius=8)
    app.ac_info_bar.pack(fill="x", padx=12, pady=(0, 6))

    app.ac_info_label = ctk.CTkLabel(
        app.ac_info_bar,
        text="Upload a file to begin. Rows with Event Names over 99 characters will appear below.",
        text_color="#888888", font=ctk.CTkFont(size=12),
    )
    app.ac_info_label.pack(side="left", padx=12, pady=6)

    # ── Column headers ────────────────────────────────────────────
    col_header = ctk.CTkFrame(tab, fg_color="#333333", height=32)
    col_header.pack(fill="x", padx=12, pady=(0, 2))
    col_header.pack_propagate(False)
    col_header.grid_columnconfigure(0, weight=3)
    col_header.grid_columnconfigure(1, weight=3)
    col_header.grid_columnconfigure(2, minsize=90)

    ctk.CTkLabel(col_header, text="Original Event Name (read-only)",
                 font=ctk.CTkFont(size=12, weight="bold"), anchor="w",
                 text_color="#cccccc").grid(row=0, column=0, padx=10, sticky="w", pady=4)
    ctk.CTkLabel(col_header, text="Edit Event Name",
                 font=ctk.CTkFont(size=12, weight="bold"), anchor="w",
                 text_color="#cccccc").grid(row=0, column=1, padx=10, sticky="w", pady=4)
    ctk.CTkLabel(col_header, text="Characters",
                 font=ctk.CTkFont(size=12, weight="bold"), anchor="center",
                 text_color="#cccccc").grid(row=0, column=2, padx=8, sticky="ew", pady=4)

    # ── Scrollable rows ───────────────────────────────────────────
    app.ac_event_scroll = ctk.CTkScrollableFrame(tab)
    app.ac_event_scroll.pack(fill="both", expand=True, padx=12, pady=4)

    # ── Bottom bar with Next button ───────────────────────────────
    bottom = ctk.CTkFrame(tab)
    bottom.pack(fill="x", padx=12, pady=8)

    app.ac_en_status = ctk.CTkLabel(
        bottom, text="", text_color="#aaaaaa", font=ctk.CTkFont(size=12)
    )
    app.ac_en_status.pack(side="left", padx=8)

    ctk.CTkButton(
        bottom, text="Next: Meeting Bodies →", width=200, height=38,
        fg_color="#2e7d32", hover_color="#1b5e20",
        font=ctk.CTkFont(size=13, weight="bold"),
        command=app._ac_next_to_meeting_bodies,
    ).pack(side="right", padx=8)

def ac_browse_file(app):
    """Let user pick the Agenda Center Excel file."""
    path = filedialog.askopenfilename(
        title="Select Agenda Center Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
    )
    if not path:
        return

    try:
        import pandas as pd
        df = pd.read_excel(path)

        # Validate required columns
        required = ["Event Name", "Event Date", "Event Time",
                    "Event Category", "Meeting Type"]
        missing_cols = [c for c in required if c not in df.columns]
        if missing_cols:
            messagebox.showerror(
                "Invalid File",
                f"Missing required columns:\n{', '.join(missing_cols)}"
            )
            return

        app.ac_source_file = path
        app.ac_df = df
        app.ac_event_name_edits = {}
        app.ac_file_label.configure(
            text=Path(path).name, text_color="#ffffff"
        )

        # Count flagged rows
        flagged = []
        for idx, row in df.iterrows():
            name = str(row.get("Event Name", "") or "")
            if len(name) > 99:
                flagged.append((idx, name))

        total = len(df)
        flag_count = len(flagged)

        if flag_count == 0:
            app.ac_info_label.configure(
                text=f"✓ All {total} rows have Event Names within 99 characters. "
                     f"No edits needed — click Next to continue.",
                text_color="#4caf50",
            )
        else:
            app.ac_info_label.configure(
                text=f"⚠  {flag_count} of {total} rows exceed 99 characters. "
                     f"Edit them below before continuing.",
                text_color="#ff9800",
            )

        app._ac_populate_event_name_rows(flagged)

        # Also pre-populate AC meeting bodies data so tab is ready
        app._ac_preload_meeting_bodies()

    except Exception as e:
        messagebox.showerror("Load Error", f"Could not load file:\n{str(e)}")

def ac_populate_event_name_rows(app, flagged: list):
    """Populate the scrollable list of flagged event name rows."""
    # Clear existing rows
    for widget in app.ac_event_scroll.winfo_children():
        widget.destroy()

    if not flagged:
        ctk.CTkLabel(
            app.ac_event_scroll,
            text="No Event Names exceed 99 characters.",
            text_color="#4caf50", font=ctk.CTkFont(size=13),
        ).pack(pady=20)
        app.ac_en_status.configure(text="No fixes needed.")
        return

    app.ac_en_status.configure(
        text=f"{len(flagged)} rows to review"
    )

    for row_idx, original_name in flagged:
        row_frame = ctk.CTkFrame(app.ac_event_scroll, fg_color="#1e1e1e", corner_radius=6)
        row_frame.pack(fill="x", pady=3, padx=2)
        row_frame.grid_columnconfigure(0, weight=3)
        row_frame.grid_columnconfigure(1, weight=3)
        row_frame.grid_columnconfigure(2, minsize=90)

        # Original (read-only label, truncated display)
        display_name = original_name if len(original_name) <= 80 else original_name[:77] + "..."
        orig_label = ctk.CTkLabel(
            row_frame, text=display_name,
            anchor="w", wraplength=380,
            text_color="#888888", font=ctk.CTkFont(size=11),
        )
        orig_label.grid(row=0, column=0, padx=10, pady=6, sticky="ew")

        # Editable entry — pre-fill with current edited value or original
        current_val = app.ac_event_name_edits.get(row_idx, original_name)
        entry_var = ctk.StringVar(value=current_val)

        # Character counter label
        char_label = ctk.CTkLabel(
            row_frame, text=f"{len(current_val)}",
            width=80, anchor="center",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#ff5252" if len(current_val) > 99 else "#4caf50",
        )
        char_label.grid(row=0, column=2, padx=8, pady=6)

        entry = ctk.CTkEntry(row_frame, textvariable=entry_var, height=32)
        entry.grid(row=0, column=1, padx=10, pady=6, sticky="ew")

        def _on_change(var=entry_var, lbl=char_label, ridx=row_idx):
            val = var.get()
            n = len(val)
            lbl.configure(
                text=str(n),
                text_color="#ff5252" if n > 99 else "#4caf50",
            )
            app.ac_event_name_edits[ridx] = val

        entry_var.trace_add("write", lambda *a, fn=_on_change: fn())

def ac_next_to_meeting_bodies(app):
    """Validate Event Names and move to AC Meeting Bodies tab."""
    if app.ac_df is None:
        messagebox.showwarning("No File", "Please upload an Excel file first.")
        return

    # Check if any flagged names are still over 99 chars
    still_over = []
    for idx, row in app.ac_df.iterrows():
        original = str(row.get("Event Name", "") or "")
        if len(original) <= 99:
            continue
        edited = app.ac_event_name_edits.get(idx, original)
        if len(edited) > 99:
            still_over.append(edited[:60] + "..." if len(edited) > 60 else edited)

    if still_over:
        messagebox.showwarning(
            "Names Still Too Long",
            f"{len(still_over)} Event Name(s) still exceed 99 characters:\n\n" +
            "\n".join(f"  • {n}" for n in still_over[:5]) +
            ("\n  ..." if len(still_over) > 5 else "") +
            "\n\nPlease fix all names before continuing."
        )
        return

    app._switch_tab("AC Meeting Bodies")

