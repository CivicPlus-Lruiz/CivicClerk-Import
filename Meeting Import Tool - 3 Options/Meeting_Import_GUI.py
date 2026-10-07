"""
Meeting Import GUI
CustomTkinter interface for processing meeting files from Excel.

Supports:
  - Single-sheet Excel (Agendas / Minutes / Videos columns)
  - Multi-tab Excel where each tab = one Meeting Body (Tab name assigned in GUI)
  - Subtypes tab: scan detected subtypes, set overrides per Body + Subtype

Drop this file next to Meeting_Import_Engine.py and run it.
"""

# ================================================================
# DEPENDENCY CHECK
# ================================================================
import sys
import subprocess
import importlib

def _ensure_deps():
    pkgs = [("customtkinter","customtkinter"),("openpyxl","openpyxl"),("pandas","pandas")]
    missing = [p for i,p in pkgs if not importlib.util.find_spec(i)]
    if missing:
        print("Installing:", missing)
        for p in missing:
            subprocess.check_call([sys.executable,"-m","pip","install",p,"--break-system-packages"],
                                  stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

if __name__ == "__main__":
    _ensure_deps()

# ================================================================
# IMPORTS
# ================================================================
import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import json
import os
import subprocess as sp
from pathlib import Path
from typing import List, Dict, Optional, Any
import pandas as pd

# ================================================================
# CONSTANTS
# ================================================================
APP_TITLE   = "Meeting Import Tool"
APP_GEO     = "1500x860"
CONFIG_FILE = Path.home() / ".meeting_import_gui.json"

# ── Filename pattern presets for Scan Subtypes ──
PATTERN_PRESETS = [
    "{Meeting Body}_{Subtype}_{File Type}_{Date}",
    "{Meeting Body}_{File Type}_{Date}",
    "{Date}_{Meeting Body}_{Subtype}_{File Type}",
    "{Date}_{Meeting Body}_{File Type}",
    "{Subtype}_{Date}_{Meeting Body}",
    "{Meeting Body}_{Subtype}_{Date}",
    "{Ignore}_{Meeting Body}_{Subtype}_{File Type}_{Date}",
    "{Ignore}_{Ignore}_{Meeting Body}_{Subtype}_{File Type}_{Date}",
    "Custom Pattern...",
]
SEPARATOR_OPTIONS = ["_", "-", " "]

# Display labels for the date-format dropdown → engine value passed to parse_filename.
# Format keys are passed directly to the engine which maps them to explicit parse rules.
# Labels use "-" as representative separator; "_" and "." work identically (engine normalises all).
DATE_FORMAT_OPTIONS = {
    "MM-DD-YYYY": "MM-DD-YYYY",
    "MM-DD-YY":   "MM-DD-YY",
    "MMDDYYYY":   "MMDDYYYY",
    "MMDDYY":     "MMDDYY",
    "DD-MM-YYYY": "DD-MM-YYYY",
    "DD-MM-YY":   "DD-MM-YY",
    "DDMMYYYY":   "DDMMYYYY",
    "DDMMYY":     "DDMMYY",
    "YYYY-MM-DD": "YYYY-MM-DD",
    "YYYYMMDD":   "YYYYMMDD",
    "YY-MM-DD":   "YY-MM-DD",
    "YYMMDD":     "YYMMDD",
}
# Reverse map: engine value → display label (identity map now that values match keys)
_DATE_FORMAT_LABEL = {v: k for k, v in DATE_FORMAT_OPTIONS.items()}

PANEL_BG    = "#2b2b2b"
ROW_A       = "#2b2b2b"
ROW_B       = "#333333"
BLUE        = "#1f538d"
BLUE_HOVER  = "#14375e"
GREEN       = "#2e7d32"
GREEN_HOVER = "#1b5e20"
RED         = "#c44536"
RED_HOVER   = "#a03529"
GOLD        = "#d4a500"
GOLD_HOVER  = "#b8900a"
GRAY        = "#555555"
GRAY_HOVER  = "#666666"
TEAL        = "#1a6b6b"
TEAL_HOVER  = "#124e4e"

# ================================================================
# DATA MODELS
# ================================================================
class TabConfig:
    """Configuration for one Excel sheet/tab = one Meeting Body."""
    def __init__(self, sheet_name: str = "", body_name: str = "",
                 meeting_type: str = "", event_category: str = "",
                 default_time: str = "00:00", date_format: str = "MM-DD-YYYY"):
        self.sheet_name     = sheet_name
        self.body_name      = body_name
        self.meeting_type   = meeting_type
        self.event_category = event_category
        self.default_time   = default_time
        self.date_format    = date_format

    def to_dict(self):
        return {k: v for k, v in vars(self).items() if not k.startswith("_")}

    @classmethod
    def from_dict(cls, d):
        return cls(**{k: d.get(k, "") for k in
                      ["sheet_name","body_name","meeting_type","event_category","default_time","date_format"]})


class SubtypeOverride:
    """Override configuration for one Body + Alias combination."""
    def __init__(self, body: str = "", alias: str = "", subtype: str = "",
                 override_event_category: str = "",
                 override_meeting_type: str = "", override_time: str = ""):
        self.body                    = body
        self.alias                   = alias    # raw token(s) from filename, comma-separated
        self.subtype                 = subtype  # display name user fills in (blank = use Event Category + Meeting)
        self.override_event_category = override_event_category
        self.override_meeting_type   = override_meeting_type
        self.override_time           = override_time

    def to_dict(self):
        return {k: v for k, v in vars(self).items() if not k.startswith("_")}

    @classmethod
    def from_dict(cls, d):
        return cls(**{k: d.get(k, "") for k in
                      ["body","alias","subtype","override_event_category",
                       "override_meeting_type","override_time"]})


# ================================================================
# DATE FORMAT OVERRIDE HELPER
# ================================================================
def _parse_with_format(fname: str, fmt_label: str):
    """
    Given a filename and a display format label (e.g. "DD-MM-YYYY  (e.g. 15-01-2025)"),
    extract the date using only that specific pattern — bypassing the engine's
    multi-pattern fallback logic.
    Returns a datetime or None.
    """
    import re
    from datetime import datetime as _dt

    fmt_key = fmt_label.split()[0].strip() if fmt_label else ""

    # Normalise filename: strip extension, replace separators with space
    base = re.sub(r"\.(pdf|mp4|mp3|docx?|srt|vtt)$", "", fname, flags=re.IGNORECASE)
    base = re.sub(r"[_\-\.]+", " ", base.lower())

    def _find_sequence(text, lengths):
        """Find first sequence of numeric tokens matching given lengths list."""
        tokens = re.findall(r"\d+", text)
        for i in range(len(tokens) - len(lengths) + 1):
            chunk = tokens[i:i+len(lengths)]
            if all(len(chunk[j]) <= lengths[j]*2 and len(chunk[j]) >= 1
                   for j in range(len(lengths))):
                return [int(x) for x in chunk]
        return None

    from Meeting_Import_Engine import pivot_year

    try:
        if fmt_key == "MM-DD-YYYY":
            seq = _find_sequence(base, [2,2,4])
            if seq: return _dt(seq[2], seq[0], seq[1])
        elif fmt_key == "DD-MM-YYYY":
            seq = _find_sequence(base, [2,2,4])
            if seq: return _dt(seq[2], seq[1], seq[0])
        elif fmt_key == "YYYY-MM-DD":
            seq = _find_sequence(base, [4,2,2])
            if seq: return _dt(seq[0], seq[1], seq[2])
        elif fmt_key == "MM-DD-YY":
            seq = _find_sequence(base, [2,2,2])
            if seq: return _dt(pivot_year(seq[2]), seq[0], seq[1])
        elif fmt_key == "DD-MM-YY":
            seq = _find_sequence(base, [2,2,2])
            if seq: return _dt(pivot_year(seq[2]), seq[1], seq[0])
        elif fmt_key == "YY-MM-DD":
            seq = _find_sequence(base, [2,2,2])
            if seq: return _dt(pivot_year(seq[0]), seq[1], seq[2])
        elif fmt_key == "MMDDYYYY":
            m = re.search(r"(\d{8})", re.sub(r"\s+","",base))
            if m:
                s = m.group(1)
                return _dt(int(s[4:]), int(s[:2]), int(s[2:4]))
        elif fmt_key == "DDMMYYYY":
            m = re.search(r"(\d{8})", re.sub(r"\s+","",base))
            if m:
                s = m.group(1)
                return _dt(int(s[4:]), int(s[2:4]), int(s[:2]))
        elif fmt_key == "YYYYMMDD":
            m = re.search(r"(\d{8})", re.sub(r"\s+","",base))
            if m:
                s = m.group(1)
                return _dt(int(s[:4]), int(s[4:6]), int(s[6:]))
        elif fmt_key == "YYMMDD":
            m = re.search(r"(\d{6})", re.sub(r"\s+","",base))
            if m:
                s = m.group(1)
                return _dt(pivot_year(int(s[:2])), int(s[2:4]), int(s[4:]))
    except (ValueError, TypeError):
        pass
    return None


# ================================================================
# URL / FILENAME HELPER
# ================================================================
def _extract_filename(raw: str) -> str:
    """
    Given a bare filename OR a full URL (https:// or ftp://),
    return just the filename with URL-encoding decoded.
    """
    from urllib.parse import urlparse, unquote
    raw = raw.strip()
    parsed = urlparse(raw)
    if parsed.scheme in ("http", "https", "ftp", "ftps"):
        fname = parsed.path.rstrip("/").split("/")[-1]
    else:
        fname = raw
    return unquote(fname)


# ================================================================
# MAIN APP
# ================================================================
class MeetingImportGUI(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry(APP_GEO)
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.input_file         = ctk.StringVar()
        self.output_file        = ctk.StringVar()
        self.tab_configs:       List[TabConfig]       = []
        self.subtype_overrides: List[SubtypeOverride] = []
        self.last_result:       Optional[Dict]        = None
        self._running           = False

        # FTP credentials (persisted in config)
        self.ftp_host     = ctk.StringVar()
        self.ftp_user     = ctk.StringVar()
        self.ftp_password = ctk.StringVar()
        self.ftp_path     = ctk.StringVar(value="/")
        self.az_sas_url   = ctk.StringVar()   # Full Azure SAS URL (not persisted)
        self.az_parsed_path = ctk.StringVar()  # Editable client path extracted from URL
        self.scan_pattern = ctk.StringVar(value=PATTERN_PRESETS[0])
        self.scan_sep     = ctk.StringVar(value="_")
        self.scan_custom  = ctk.StringVar()
        self.scan_filter       = ctk.StringVar(value="All Meeting Bodies")
        self.ftp_structure     = ctk.StringVar(value="None")
        self.az_structure      = ctk.StringVar(value="None")
        self.flat_pattern      = ctk.StringVar(value=PATTERN_PRESETS[0])
        self.flat_sep          = ctk.StringVar(value="_")
        self.flat_custom       = ctk.StringVar()

        # Error fix dicts — persisted to _Config.json
        # date_fixes:  {filename: "MM/DD/YYYY"}
        # dup_fixes:   {(fname1, fname2, file_field): "keep1"|"keep2"|"other1"|"other2"}
        self.date_fixes: Dict[str, str]               = {}
        self.dup_fixes:  Dict[tuple, str]             = {}
        self.body_fixes: Dict[str, str]               = {}  # {filename: sheet_name}
        self._pending_errors: List[dict]              = []  # raw errors from last run

        self._build_ui()
        self._load_config()

    # ================================================================
    # UI CONSTRUCTION
    # ================================================================
    def _build_ui(self):
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=10, pady=10)

        tab_bar = ctk.CTkFrame(main, fg_color=PANEL_BG, height=60)
        tab_bar.pack(fill="x", pady=(0,10))
        tab_bar.pack_propagate(False)

        btn_container = ctk.CTkFrame(tab_bar, fg_color="transparent")
        btn_container.place(relx=0.5, rely=0.5, anchor="center")

        self.tabs = ["Setup", "Sheet Mapping", "Subtypes", "Run", "Errors", "Results"]
        self.tab_btns:   Dict[str, ctk.CTkButton] = {}
        self.tab_frames: Dict[str, ctk.CTkFrame]  = {}

        for i, name in enumerate(self.tabs):
            btn = ctk.CTkButton(
                btn_container, text=name, width=140, height=40,
                font=ctk.CTkFont(size=14, weight="bold"),
                fg_color=BLUE if i == 0 else "#3a3a3a",
                hover_color=BLUE_HOVER if i == 0 else "#4a4a4a",
                corner_radius=8,
                command=lambda t=name: self._switch_tab(t)
            )
            btn.pack(side="left", padx=5)
            self.tab_btns[name] = btn

        self.content = ctk.CTkFrame(main)
        self.content.pack(fill="both", expand=True)

        for name in self.tabs:
            self.tab_frames[name] = ctk.CTkFrame(self.content)

        self._build_setup_tab()
        self._build_mapping_tab()
        self._build_subtypes_tab()
        self._build_run_tab()
        self._build_errors_tab()
        self._build_results_tab()

        self.current_tab = "Setup"
        self.tab_frames["Setup"].pack(fill="both", expand=True)

    def _switch_tab(self, name: str):
        if self.current_tab in self.tab_frames:
            self.tab_frames[self.current_tab].pack_forget()
        for n, b in self.tab_btns.items():
            b.configure(fg_color=BLUE if n == name else "#3a3a3a",
                        hover_color=BLUE_HOVER if n == name else "#4a4a4a")
        self.tab_frames[name].pack(fill="both", expand=True)
        self.current_tab = name

    # ================================================================
    # TAB 1 — SETUP
    # ================================================================
    def _build_setup_tab(self):
        tab = self.tab_frames["Setup"]

        # Header + Next button row (always visible, above scroll)
        top = ctk.CTkFrame(tab, fg_color="transparent")
        top.pack(fill="x", padx=20, pady=(18,4))
        ctk.CTkLabel(top, text="Setup",
                     font=ctk.CTkFont(size=20, weight="bold")).pack(side="left")

        ctk.CTkLabel(tab, text="Choose your data source: an existing Excel file, or scan directly from FTP or Azure.",
                     text_color="gray").pack(anchor="w", padx=20, pady=(0,6))

        # Scrollable body
        setup_scroll = ctk.CTkScrollableFrame(tab)
        setup_scroll.pack(fill="both", expand=True, padx=0, pady=0)
        tab = setup_scroll  # redirect all further .pack() calls into the scroll frame

        # ── Source selector toggle ──
        src_bar = ctk.CTkFrame(tab, fg_color=PANEL_BG, corner_radius=10)
        src_bar.pack(fill="x", padx=20, pady=(0,10))

        src_inner = ctk.CTkFrame(src_bar, fg_color="transparent")
        src_inner.pack(padx=14, pady=12, anchor="w")

        ctk.CTkLabel(src_inner, text="Data Source:",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(side="left", padx=(0,14))

        self._source_mode = ctk.StringVar(value="excel")

        def _toggle_source(mode):
            self._source_mode.set(mode)
            excel_sec.pack_forget()
            ftp_sec.pack_forget()
            az_sec.pack_forget()
            for btn, m in [(src_excel_btn,"excel"),(src_ftp_btn,"ftp"),(src_az_btn,"azure")]:
                btn.configure(fg_color=BLUE if m==mode else "#3a3a3a",
                              hover_color=BLUE_HOVER if m==mode else "#4a4a4a")
            if mode == "excel":
                excel_sec.pack(fill="x", padx=20, pady=4)
            elif mode == "ftp":
                ftp_sec.pack(fill="x", padx=20, pady=4)
            else:
                az_sec.pack(fill="x", padx=20, pady=4)

        src_excel_btn = ctk.CTkButton(src_inner, text="📄  Excel File", width=150, height=34,
                                      fg_color=BLUE, hover_color=BLUE_HOVER,
                                      command=lambda: _toggle_source("excel"))
        src_excel_btn.pack(side="left", padx=(0,6))

        src_ftp_btn = ctk.CTkButton(src_inner, text="🌐  FTP Server", width=150, height=34,
                                    fg_color="#3a3a3a", hover_color="#4a4a4a",
                                    command=lambda: _toggle_source("ftp"))
        src_ftp_btn.pack(side="left", padx=(0,6))

        src_az_btn = ctk.CTkButton(src_inner, text="☁  Azure Storage", width=160, height=34,
                                   fg_color="#3a3a3a", hover_color="#4a4a4a",
                                   command=lambda: _toggle_source("azure"))
        src_az_btn.pack(side="left")

        # ── Excel panel ──
        excel_sec = ctk.CTkFrame(tab, fg_color=PANEL_BG, corner_radius=10)
        excel_sec.pack(fill="x", padx=20, pady=4)

        ctk.CTkLabel(excel_sec, text="Input Excel File",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=14, pady=(12,4))
        ctk.CTkLabel(excel_sec,
                     text="Spreadsheet with Agendas / Minutes / Videos / Packets columns.\n"
                          "Single sheet = auto-detect bodies.  Multiple sheets = one body per sheet.",
                     text_color="gray", font=ctk.CTkFont(size=11)
                     ).pack(anchor="w", padx=14, pady=(0,8))
        xrow = ctk.CTkFrame(excel_sec, fg_color="transparent")
        xrow.pack(fill="x", padx=14, pady=(0,14))
        ctk.CTkEntry(xrow, textvariable=self.input_file, width=660,
                     placeholder_text="Browse or type path to .xlsx file"
                     ).pack(side="left", padx=(0,8))
        ctk.CTkButton(xrow, text="Browse…", width=110,
                      command=self._browse_input).pack(side="left", padx=(0,8))
        ctk.CTkButton(xrow, text="↻ Scan Sheets", width=130,
                      fg_color=GRAY, hover_color=GRAY_HOVER,
                      command=self._scan_sheets).pack(side="left")

        # ── FTP panel ──
        ftp_sec = ctk.CTkFrame(tab, fg_color=PANEL_BG, corner_radius=10)
        # (not packed yet — shown when user selects FTP)

        ctk.CTkLabel(ftp_sec, text="FTP Connection",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=14, pady=(12,4))

        # Folder Structure selector — shown first, rest appears after selection
        ftp_struct_row = ctk.CTkFrame(ftp_sec, fg_color="transparent")
        ftp_struct_row.pack(fill="x", padx=14, pady=(0,8))
        ctk.CTkLabel(ftp_struct_row, text="Folder Structure:", anchor="e", width=130
                     ).pack(side="left", padx=(0,8))

        ftp_creds_frame = ctk.CTkFrame(ftp_sec, fg_color="transparent")
        ftp_flat_frame  = ctk.CTkFrame(ftp_sec, fg_color="transparent")
        ftp_conn_frame  = ctk.CTkFrame(ftp_sec, fg_color="transparent")

        def _ftp_struct_changed(v):
            self.ftp_structure.set(v)
            ftp_creds_frame.pack_forget()
            ftp_flat_frame.pack_forget()
            ftp_conn_frame.pack_forget()
            if v == "None":
                return
            ftp_creds_frame.pack(fill="x", padx=14, pady=(0,4))
            if v == "No Folder Structure":
                ftp_flat_frame.pack(fill="x", padx=14, pady=(0,4))
            ftp_conn_frame.pack(fill="x", padx=14, pady=(4,14))

        ctk.CTkOptionMenu(ftp_struct_row, variable=self.ftp_structure, width=220,
                          values=["None", "Structured Folders", "No Folder Structure"],
                          command=_ftp_struct_changed
                          ).pack(side="left")

        # Credentials (inside ftp_creds_frame)
        ctk.CTkLabel(ftp_creds_frame,
                     text="Connect to the FTP server to scan the directory structure and build the input Excel automatically.",
                     text_color="gray", font=ctk.CTkFont(size=11)
                     ).pack(anchor="w", pady=(4,6))
        cred_grid = ctk.CTkFrame(ftp_creds_frame, fg_color="transparent")
        cred_grid.pack(fill="x")
        ctk.CTkLabel(cred_grid, text="Host / Address:", anchor="e", width=120
                     ).grid(row=0, column=0, padx=(0,8), pady=5, sticky="e")
        ctk.CTkEntry(cred_grid, textvariable=self.ftp_host, width=380,
                     placeholder_text="e.g. imports.civicclerk.com"
                     ).grid(row=0, column=1, padx=(0,20), pady=5, sticky="w")
        ctk.CTkLabel(cred_grid, text="Remote Path:", anchor="e", width=110
                     ).grid(row=0, column=2, padx=(0,8), pady=5, sticky="e")
        ctk.CTkEntry(cred_grid, textvariable=self.ftp_path, width=180,
                     placeholder_text="/"
                     ).grid(row=0, column=3, pady=5, sticky="w")
        ctk.CTkLabel(cred_grid, text="Username:", anchor="e", width=120
                     ).grid(row=1, column=0, padx=(0,8), pady=5, sticky="e")
        ctk.CTkEntry(cred_grid, textvariable=self.ftp_user, width=380,
                     placeholder_text="e.g. NASSAUCOFL"
                     ).grid(row=1, column=1, padx=(0,20), pady=5, sticky="w")
        ctk.CTkLabel(cred_grid, text="Password:", anchor="e", width=110
                     ).grid(row=1, column=2, padx=(0,8), pady=5, sticky="e")
        ctk.CTkEntry(cred_grid, textvariable=self.ftp_password, width=180, show="•"
                     ).grid(row=1, column=3, pady=5, sticky="w")

        # Flat pattern row
        self._build_flat_pattern_row(ftp_flat_frame)

        # Connect row
        self.ftp_connect_btn = ctk.CTkButton(
            ftp_conn_frame, text="\U0001f50c  Connect & Scan FTP", width=200, height=38,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color=TEAL, hover_color=TEAL_HOVER,
            command=self._ftp_scan)
        self.ftp_connect_btn.pack(side="left", padx=(0,16))
        self.ftp_status_label = ctk.CTkLabel(
            ftp_conn_frame, text="Not connected.", text_color="gray",
            font=ctk.CTkFont(size=12))
        self.ftp_status_label.pack(side="left")

        # ── Azure Blob Storage panel ──
        az_sec = ctk.CTkFrame(tab, fg_color=PANEL_BG, corner_radius=10)
        # (not packed yet — shown when user selects Azure)

        ctk.CTkLabel(az_sec, text="Azure Blob Storage",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=14, pady=(12,4))

        # Folder Structure selector — shown first
        az_struct_row = ctk.CTkFrame(az_sec, fg_color="transparent")
        az_struct_row.pack(fill="x", padx=14, pady=(0,8))
        ctk.CTkLabel(az_struct_row, text="Folder Structure:", anchor="e", width=130
                     ).pack(side="left", padx=(0,8))

        az_creds_frame = ctk.CTkFrame(az_sec, fg_color="transparent")
        az_flat_frame  = ctk.CTkFrame(az_sec, fg_color="transparent")
        az_conn_frame  = ctk.CTkFrame(az_sec, fg_color="transparent")

        def _az_struct_changed(v):
            self.az_structure.set(v)
            az_creds_frame.pack_forget()
            az_flat_frame.pack_forget()
            az_conn_frame.pack_forget()
            if v == "None":
                return
            az_creds_frame.pack(fill="x", padx=14, pady=(0,4))
            if v == "No Folder Structure":
                az_flat_frame.pack(fill="x", padx=14, pady=(0,4))
            az_conn_frame.pack(fill="x", padx=14, pady=(4,14))

        ctk.CTkOptionMenu(az_struct_row, variable=self.az_structure, width=220,
                          values=["None", "Structured Folders", "No Folder Structure"],
                          command=_az_struct_changed
                          ).pack(side="left")

        # Credentials (inside az_creds_frame)
        ctk.CTkLabel(az_creds_frame,
                     text="Paste a full Azure SAS URL. The client path is extracted automatically.",
                     text_color="gray", font=ctk.CTkFont(size=11)
                     ).pack(anchor="w", pady=(4,6))
        az_url_row = ctk.CTkFrame(az_creds_frame, fg_color="transparent")
        az_url_row.pack(fill="x", pady=(0,4))
        ctk.CTkLabel(az_url_row, text="SAS URL:", anchor="e", width=90
                     ).pack(side="left", padx=(0,8))
        ctk.CTkEntry(az_url_row, textvariable=self.az_sas_url, width=660,
                     placeholder_text="https://account.blob.core.windows.net/container/path/...?sv=...&sig=..."
                     ).pack(side="left")

        # Flat pattern row
        self._build_flat_pattern_row(az_flat_frame)

        # Connect row
        self.az_connect_btn = ctk.CTkButton(
            az_conn_frame, text="\u2601  Connect & Scan Azure", width=210, height=38,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color=TEAL, hover_color=TEAL_HOVER,
            command=self._azure_scan)
        self.az_connect_btn.pack(side="left", padx=(0,16))
        self.az_status_label = ctk.CTkLabel(
            az_conn_frame, text="Not connected.", text_color="gray",
            font=ctk.CTkFont(size=12))
        self.az_status_label.pack(side="left")

        # ── Output file (always visible) ──
        sec2 = ctk.CTkFrame(tab, fg_color=PANEL_BG, corner_radius=10)
        sec2.pack(fill="x", padx=20, pady=8)
        ctk.CTkLabel(sec2, text="Output Excel File",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=14, pady=(12,4))
        ctk.CTkLabel(sec2, text="Where the final CivicClerk import file will be saved.",
                     text_color="gray", font=ctk.CTkFont(size=11)
                     ).pack(anchor="w", padx=14, pady=(0,8))
        row2 = ctk.CTkFrame(sec2, fg_color="transparent")
        row2.pack(fill="x", padx=14, pady=(0,14))
        ctk.CTkEntry(row2, textvariable=self.output_file, width=680,
                     placeholder_text="Browse or type output path"
                     ).pack(side="left", padx=(0,8))
        ctk.CTkButton(row2, text="Browse…", width=110,
                      command=self._browse_output).pack(side="left")

        # ── Info box ──
        info = ctk.CTkFrame(tab, fg_color="#1a3a1a", corner_radius=10)
        info.pack(fill="x", padx=20, pady=(10,4))
        ctk.CTkLabel(info,
                     text="ℹ  Workflow\n\n"
                          "Option A — Excel:        Browse to your existing input spreadsheet.  Sheets are scanned automatically.\n"
                          "Option B — FTP:          Enter credentials and click Connect & Scan FTP.\n"
                          "Option C — Azure:        Enter client path + SAS token and click Connect & Scan Azure.\n\n"
                          "Options B & C walk the folder structure (Body → Agendas/Minutes/Videos/Packets), build the input Excel automatically, and pre-populate Sheet Mapping.",
                     anchor="w", justify="left", text_color="#aaffaa",
                     font=ctk.CTkFont(size=12)
                     ).pack(padx=16, pady=14, anchor="w")

        ctk.CTkButton(tab, text="Next: Sheet Mapping  →", width=220,
                      fg_color=BLUE, hover_color=BLUE_HOVER,
                      command=lambda: self._switch_tab("Sheet Mapping")
                      ).pack(anchor="e", padx=20, pady=(4,14))



    # ================================================================
    # TAB 2 — SHEET MAPPING
    # ================================================================
    def _build_mapping_tab(self):
        tab = self.tab_frames["Sheet Mapping"]

        hdr = ctk.CTkFrame(tab, fg_color="transparent")
        hdr.pack(fill="x", padx=20, pady=(16,4))
        ctk.CTkLabel(hdr, text="Sheet Mapping",
                     font=ctk.CTkFont(size=20, weight="bold")).pack(side="left")

        btn_r = ctk.CTkFrame(hdr, fg_color="transparent")
        btn_r.pack(side="right")
        ctk.CTkButton(btn_r, text="+ Add Row", width=110,
                      command=self._add_mapping_row).pack(side="left", padx=4)
        ctk.CTkButton(btn_r, text="Remove Selected", width=140,
                      fg_color=RED, hover_color=RED_HOVER,
                      command=self._remove_selected_rows).pack(side="left", padx=4)
        ctk.CTkButton(btn_r, text="↻ Scan Sheets", width=130,
                      fg_color=GRAY, hover_color=GRAY_HOVER,
                      command=self._scan_sheets).pack(side="left", padx=4)
        ctk.CTkButton(btn_r, text="📂 Scan Folder", width=130,
                      fg_color=GRAY, hover_color=GRAY_HOVER,
                      command=self._scan_output_folder).pack(side="left", padx=4)
        ctk.CTkButton(btn_r, text="⬇ Load Config", width=130,
                      fg_color="#1a5c3a", hover_color="#1e7a4e",
                      command=self._load_customer_config).pack(side="left", padx=4)

        ctk.CTkLabel(tab,
                     text="Each row maps one Excel sheet to a Meeting Body. "
                          "Leave Body Name blank to use auto-detection from filenames.",
                     text_color="gray", font=ctk.CTkFont(size=11)
                     ).pack(anchor="w", padx=20, pady=(0,10))

        col_hdr = ctk.CTkFrame(tab, fg_color=PANEL_BG, height=34)
        col_hdr.pack(fill="x", padx=20)
        col_hdr.pack_propagate(False)
        col_hdr.grid_columnconfigure(0, minsize=28)
        col_hdr.grid_columnconfigure(1, weight=1, minsize=160)
        col_hdr.grid_columnconfigure(2, weight=2, minsize=220)
        col_hdr.grid_columnconfigure(3, weight=1, minsize=160)
        col_hdr.grid_columnconfigure(4, weight=1, minsize=160)
        col_hdr.grid_columnconfigure(5, minsize=130)
        col_hdr.grid_columnconfigure(6, minsize=110)
        col_hdr.grid_columnconfigure(7, minsize=16)  # spacer matching scrollbar width

        for col, text in enumerate(["☐","Sheet / Tab Name",
                                     "Body Name",
                                     "Meeting Type","Event Category","Default Time","Date Format"]):
            ctk.CTkLabel(col_hdr, text=text, font=ctk.CTkFont(weight="bold"),
                         anchor="w").grid(row=0, column=col, padx=(4, 0), pady=6, sticky="w")

        self.mapping_scroll = ctk.CTkScrollableFrame(tab)
        self.mapping_scroll.pack(fill="both", expand=True, padx=20, pady=(0,6))
        self._refresh_mapping_display()

        ctk.CTkButton(tab, text="Next: Subtypes  →", width=180,
                      fg_color=BLUE, hover_color=BLUE_HOVER,
                      command=lambda: self._switch_tab("Subtypes")
                      ).pack(anchor="e", padx=20, pady=8)

    def _refresh_mapping_display(self):
        for w in self.mapping_scroll.winfo_children():
            w.destroy()

        for idx, tc in enumerate(self.tab_configs):
            color = ROW_A if idx % 2 == 0 else ROW_B
            row = ctk.CTkFrame(self.mapping_scroll, fg_color=color)
            row.pack(fill="x", pady=2)
            row.grid_columnconfigure(0, minsize=28)
            row.grid_columnconfigure(1, weight=1, minsize=160)
            row.grid_columnconfigure(2, weight=2, minsize=220)
            row.grid_columnconfigure(3, weight=1, minsize=160)
            row.grid_columnconfigure(4, weight=1, minsize=160)
            row.grid_columnconfigure(5, minsize=130)
            row.grid_columnconfigure(6, minsize=110)

            sel_var = ctk.BooleanVar()
            ctk.CTkCheckBox(row, text="", variable=sel_var, width=28,
                            command=lambda i=idx, v=sel_var:
                            setattr(self.tab_configs[i], '_sel', v.get())
                            ).grid(row=0, column=0, padx=4, pady=6, sticky="w")

            for col_idx, (attr, placeholder) in enumerate([
                ("sheet_name",     "Tab name in Excel"),
                ("body_name",      "e.g. Board Of County Commissioners"),
                ("meeting_type",   "e.g. Regular"),
                ("event_category", "e.g. Board Meeting"),
            ], start=1):
                var = ctk.StringVar(value=getattr(tc, attr))
                ctk.CTkEntry(row, textvariable=var, placeholder_text=placeholder
                             ).grid(row=0, column=col_idx, padx=4, pady=6, sticky="ew")
                def _on_mapping_change(*a, i=idx, v=var, at=attr):
                    setattr(self.tab_configs[i], at, v.get())
                    if at in ("meeting_type", "event_category", "body_name"):
                        self._auto_save_customer_config()
                var.trace_add("write", _on_mapping_change)

            time_f = ctk.CTkFrame(row, fg_color="transparent")
            time_f.grid(row=0, column=5, padx=4, pady=6, sticky="w")
            parts = tc.default_time.split(":") if ":" in tc.default_time else ["00","00"]
            h_var = ctk.StringVar(value=parts[0].zfill(2))
            m_var = ctk.StringVar(value=(parts[1].zfill(2) if len(parts) > 1 else "00"))

            def _upd(val=None, i=idx, hv=h_var, mv=m_var):
                self.tab_configs[i].default_time = f"{hv.get()}:{mv.get()}"
                self._auto_save_customer_config()

            ctk.CTkOptionMenu(time_f, variable=h_var, width=62,
                              values=[f"{h:02d}" for h in range(24)],
                              command=_upd).pack(side="left", padx=1)
            ctk.CTkLabel(time_f, text=":", font=ctk.CTkFont(size=13, weight="bold")).pack(side="left")
            ctk.CTkOptionMenu(time_f, variable=m_var, width=62,
                              values=["00","15","30","45"],
                              command=_upd).pack(side="left", padx=1)
            tc._h_var = h_var
            tc._m_var = m_var

            # Date Format dropdown — show friendly labels, store engine value
            _df_labels  = list(DATE_FORMAT_OPTIONS.keys())
            _saved_df   = tc.date_format or "MM-DD-YYYY"
            _df_display = _DATE_FORMAT_LABEL.get(_saved_df, _df_labels[0])
            df_var = ctk.StringVar(value=_df_display)
            def _on_df_change(label, i=idx):
                engine_val = DATE_FORMAT_OPTIONS.get(label, "MM-DD-YYYY")
                setattr(self.tab_configs[i], "date_format", engine_val)
                self._auto_save_customer_config()
            ctk.CTkOptionMenu(row, variable=df_var, width=130,
                              values=_df_labels,
                              command=_on_df_change,
                              ).grid(row=0, column=6, padx=4, pady=6, sticky="w")
            tc._df_var = df_var

    def _add_mapping_row(self):
        self.tab_configs.append(TabConfig())
        self._refresh_mapping_display()

    def _remove_selected_rows(self):
        self.tab_configs = [tc for tc in self.tab_configs
                            if not getattr(tc, '_sel', False)]
        self._refresh_mapping_display()

    # ================================================================
    # TAB 3 — SUBTYPES
    # ================================================================
    def _build_subtypes_tab(self):
        tab = self.tab_frames["Subtypes"]

        hdr = ctk.CTkFrame(tab, fg_color="transparent")
        hdr.pack(fill="x", padx=20, pady=(16,4))
        ctk.CTkLabel(hdr, text="Subtypes & Overrides",
                     font=ctk.CTkFont(size=20, weight="bold")).pack(side="left")

        btn_r = ctk.CTkFrame(hdr, fg_color="transparent")
        btn_r.pack(side="right")
        ctk.CTkButton(btn_r, text="🔍  Scan Subtypes", width=155,
                      fg_color=TEAL, hover_color=TEAL_HOVER,
                      command=self._scan_subtypes).pack(side="left", padx=4)
        ctk.CTkButton(btn_r, text="+ Add Row", width=110,
                      command=self._add_subtype_row).pack(side="left", padx=4)
        ctk.CTkButton(btn_r, text="Remove Selected", width=140,
                      fg_color=RED, hover_color=RED_HOVER,
                      command=self._remove_selected_subtypes).pack(side="left", padx=4)
        ctk.CTkButton(btn_r, text="Clear All", width=100,
                      fg_color=GRAY, hover_color=GRAY_HOVER,
                      command=self._clear_subtypes).pack(side="left", padx=4)

        # ── Filter + Pattern toolbar ──
        toolbar = ctk.CTkFrame(tab, fg_color=PANEL_BG, corner_radius=8)
        toolbar.pack(fill="x", padx=20, pady=(4,4))

        # Row 1: Filter by body
        frow = ctk.CTkFrame(toolbar, fg_color="transparent")
        frow.pack(fill="x", padx=10, pady=(8,4))
        ctk.CTkLabel(frow, text="Filter by Body:", width=110, anchor="e"
                     ).pack(side="left", padx=(0,6))
        self.filter_menu = ctk.CTkOptionMenu(
            frow, variable=self.scan_filter, width=260,
            values=["All Meeting Bodies"],
            command=lambda v: self._refresh_subtypes_display())
        self.filter_menu.pack(side="left", padx=(0,20))

        # Row 2: Pattern + separator + custom + scan button
        prow = ctk.CTkFrame(toolbar, fg_color="transparent")
        prow.pack(fill="x", padx=10, pady=(0,8))
        ctk.CTkLabel(prow, text="Pattern:", width=110, anchor="e"
                     ).pack(side="left", padx=(0,6))
        self.pattern_menu = ctk.CTkOptionMenu(
            prow, variable=self.scan_pattern, width=340,
            values=PATTERN_PRESETS,
            command=self._on_pattern_select)
        self.pattern_menu.pack(side="left", padx=(0,8))
        ctk.CTkLabel(prow, text="Sep:", width=34, anchor="e"
                     ).pack(side="left", padx=(0,4))
        ctk.CTkOptionMenu(prow, variable=self.scan_sep, width=60,
                          values=SEPARATOR_OPTIONS).pack(side="left", padx=(0,8))
        self.custom_pattern_entry = ctk.CTkEntry(
            prow, textvariable=self.scan_custom, width=240,
            placeholder_text="Custom pattern e.g. {Ignore}_{Meeting Body}_{Subtype}_{Date}")
        self.custom_pattern_entry.pack(side="left", padx=(0,8))
        self.custom_pattern_entry.configure(state="disabled")

        info = ctk.CTkFrame(tab, fg_color="#1a2a3a", corner_radius=8)
        info.pack(fill="x", padx=20, pady=(0,4))

        info_cols = ctk.CTkFrame(info, fg_color="transparent")
        info_cols.pack(fill="x", padx=12, pady=8)

        right = ctk.CTkFrame(info_cols, fg_color="transparent")
        right.pack(side="left", anchor="nw")
        ctk.CTkLabel(right, text="Pattern Token Key",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color="#aaccff", anchor="w").pack(anchor="w")
        ctk.CTkLabel(right,
                     text="{Meeting Body}  -> matches body name (exact, spaced, or compact)\n"
                          "{Subtype}       -> the segment captured as the alias token\n"
                          "{File Type}     -> matches Agenda, Minutes, Video, Packet, Other\n"
                          "{Date}          -> matches date segments (YYYY MM DD or full date)\n"
                          "{Ignore}        -> matches any single segment, discards it (repeatable)",
                     anchor="w", justify="left", text_color="#aaccff",
                     font=ctk.CTkFont(size=11)).pack(anchor="w")

        col_hdr = ctk.CTkFrame(tab, fg_color=PANEL_BG, height=34)
        col_hdr.pack(fill="x", padx=20)
        col_hdr.pack_propagate(False)
        col_hdr.grid_columnconfigure(0, minsize=28)            # checkbox
        col_hdr.grid_columnconfigure(1, minsize=52)            # up/down arrows
        col_hdr.grid_columnconfigure(2, weight=2, minsize=180) # body
        col_hdr.grid_columnconfigure(3, weight=2, minsize=180) # alias
        col_hdr.grid_columnconfigure(4, weight=1, minsize=130) # subtype
        col_hdr.grid_columnconfigure(5, weight=2, minsize=180) # override event category
        col_hdr.grid_columnconfigure(6, weight=1, minsize=140) # override meeting type
        col_hdr.grid_columnconfigure(7, minsize=160)           # override time

        for col, text in enumerate(["☐", "▲▼", "Body", "Alias (search term)",
                                     "Subtype (output name)",
                                     "Override Event Category",
                                     "Override Meeting Type", "Override Time"]):
            ctk.CTkLabel(col_hdr, text=text, font=ctk.CTkFont(weight="bold"),
                         anchor="w").grid(row=0, column=col, padx=6, pady=6, sticky="ew")

        self.subtypes_scroll = ctk.CTkScrollableFrame(tab)
        self.subtypes_scroll.pack(fill="both", expand=True, padx=20, pady=(0,6))
        self._refresh_subtypes_display()

        ctk.CTkButton(tab, text="Next: Run  →", width=160,
                      fg_color=BLUE, hover_color=BLUE_HOVER,
                      command=lambda: self._switch_tab("Run")
                      ).pack(anchor="e", padx=20, pady=8)

    def _move_subtype_row(self, idx: int, direction: int):
        """Move a subtype row up (-1) or down (+1), mark as manually ordered."""
        lst = self.subtype_overrides
        new_idx = idx + direction
        if 0 <= new_idx < len(lst):
            lst[idx], lst[new_idx] = lst[new_idx], lst[idx]
            self._subtypes_manually_ordered = True
            self._refresh_subtypes_display()

    def _refresh_subtypes_display(self):
        for w in self.subtypes_scroll.winfo_children():
            w.destroy()

        # Update filter dropdown options from current sheet names
        body_opts = ["All Meeting Bodies"] + [
            tc.sheet_name for tc in self.tab_configs if tc.sheet_name.strip()]
        if hasattr(self, "filter_menu"):
            self.filter_menu.configure(values=body_opts)
            if self.scan_filter.get() not in body_opts:
                self.scan_filter.set("All Meeting Bodies")

        filt = self.scan_filter.get() if hasattr(self, "scan_filter") else "All Meeting Bodies"
        visible = [so for so in self.subtype_overrides
                   if filt == "All Meeting Bodies" or so.body == filt]
        total = len(visible)
        for idx, so in enumerate(visible):
            color = ROW_A if idx % 2 == 0 else ROW_B
            row = ctk.CTkFrame(self.subtypes_scroll, fg_color=color)
            row.pack(fill="x", pady=2)
            row.grid_columnconfigure(0, minsize=28)             # checkbox
            row.grid_columnconfigure(1, minsize=52)             # up/down
            row.grid_columnconfigure(2, weight=2, minsize=180)  # body
            row.grid_columnconfigure(3, weight=2, minsize=180)  # alias
            row.grid_columnconfigure(4, weight=1, minsize=130)  # subtype
            row.grid_columnconfigure(5, weight=2, minsize=180)  # override event category
            row.grid_columnconfigure(6, weight=1, minsize=140)  # override meeting type
            row.grid_columnconfigure(7, minsize=160)            # override time

            g_idx = self.subtype_overrides.index(so)

            sel_var = ctk.BooleanVar()
            ctk.CTkCheckBox(row, text="", variable=sel_var, width=28,
                            command=lambda i=g_idx, v=sel_var:
                            setattr(self.subtype_overrides[i], "_sel", v.get())
                            ).grid(row=0, column=0, padx=4, pady=5, sticky="w")

            # Up / Down arrows
            arrow_f = ctk.CTkFrame(row, fg_color="transparent")
            arrow_f.grid(row=0, column=1, padx=2, pady=2, sticky="w")
            ctk.CTkButton(arrow_f, text="▲", width=22, height=20,
                          fg_color=GRAY, hover_color=GRAY_HOVER,
                          font=ctk.CTkFont(size=10),
                          state="normal" if g_idx > 0 else "disabled",
                          command=lambda i=g_idx: self._move_subtype_row(i, -1)
                          ).pack(side="top", pady=1)
            ctk.CTkButton(arrow_f, text="▼", width=22, height=20,
                          fg_color=GRAY, hover_color=GRAY_HOVER,
                          font=ctk.CTkFont(size=10),
                          state="normal" if g_idx < len(self.subtype_overrides) - 1 else "disabled",
                          command=lambda i=g_idx: self._move_subtype_row(i, 1)
                          ).pack(side="top", pady=1)

            # Body — dropdown from Sheet Mapping sheet names
            body_names = [""] + [tc.sheet_name for tc in self.tab_configs if tc.sheet_name.strip()]
            body_var = ctk.StringVar(value=so.body)
            ctk.CTkOptionMenu(row, variable=body_var,
                              values=body_names,
                              fg_color="#222222", text_color="#cccccc",
                              button_color="#333333", button_hover_color="#444444",
                              dropdown_fg_color="#222222", dropdown_text_color="#cccccc"
                              ).grid(row=0, column=2, padx=4, pady=5, sticky="ew")
            body_var.trace_add("write", lambda *a, i=idx, v=body_var:
                               setattr(self.subtype_overrides[i], "body", v.get()))

            # Alias — amber tint, auto-populated by scan, user can edit
            alias_var = ctk.StringVar(value=so.alias)
            ctk.CTkEntry(row, textvariable=alias_var,
                         fg_color="#2a2200", text_color="#ffcc66",
                         placeholder_text="e.g. SpecialMeeting"
                         ).grid(row=0, column=3, padx=4, pady=5, sticky="ew")
            alias_var.trace_add("write", lambda *a, i=g_idx, v=alias_var:
                                setattr(self.subtype_overrides[i], "alias", v.get()))

            # Subtype — blue tint, user fills in
            sub_var = ctk.StringVar(value=so.subtype)
            ctk.CTkEntry(row, textvariable=sub_var,
                         fg_color="#222233", text_color="#aaaaff",
                         placeholder_text="e.g. Special Meeting"
                         ).grid(row=0, column=4, padx=4, pady=5, sticky="ew")
            sub_var.trace_add("write", lambda *a, i=g_idx, v=sub_var:
                              setattr(self.subtype_overrides[i], "subtype", v.get()))

            # Override fields
            for col_idx, (attr, placeholder) in enumerate([
                ("override_event_category", "e.g. Board Of County Commissioners"),
                ("override_meeting_type",   "e.g. Special Meeting"),
            ], start=5):
                var = ctk.StringVar(value=getattr(so, attr))
                ctk.CTkEntry(row, textvariable=var, placeholder_text=placeholder
                             ).grid(row=0, column=col_idx, padx=4, pady=5, sticky="ew")
                var.trace_add("write", lambda *a, i=g_idx, v=var, at=attr:
                              setattr(self.subtype_overrides[i], at, v.get()))

            # Override time picker
            time_f = ctk.CTkFrame(row, fg_color="transparent")
            time_f.grid(row=0, column=7, padx=4, pady=5, sticky="w")
            parts = so.override_time.split(":") if ":" in so.override_time else ["",""]
            h_val = parts[0].zfill(2) if parts[0].strip() else "(none)"
            m_val = parts[1].zfill(2) if len(parts) > 1 and parts[1].strip() else "00"
            h_var = ctk.StringVar(value=h_val)
            m_var = ctk.StringVar(value=m_val)

            def _upd_time(val=None, i=g_idx, hv=h_var, mv=m_var):
                h = hv.get()
                self.subtype_overrides[i].override_time = (
                    f"{h}:{mv.get()}" if h != "(none)" else "")

            ctk.CTkOptionMenu(time_f, variable=h_var, width=80,
                              values=["(none)"] + [f"{h:02d}" for h in range(24)],
                              command=_upd_time).pack(side="left", padx=1)
            ctk.CTkLabel(time_f, text=":", font=ctk.CTkFont(size=13, weight="bold")
                         ).pack(side="left")
            ctk.CTkOptionMenu(time_f, variable=m_var, width=62,
                              values=["00","15","30","45"],
                              command=_upd_time).pack(side="left", padx=1)
            so._h_var = h_var
            so._m_var = m_var

    def _on_pattern_select(self, value):
        """Enable/disable custom pattern entry based on selection."""
        if value == "Custom Pattern...":
            self.custom_pattern_entry.configure(state="normal")
        else:
            self.custom_pattern_entry.configure(state="disabled")

    def _get_active_pattern(self):
        """Return (pattern_str, separator) to use for scanning."""
        pat = self.scan_pattern.get()
        sep = self.scan_sep.get()
        if pat == "Custom Pattern...":
            pat = self.scan_custom.get().strip()
        return pat, sep

    def _extract_alias_from_pattern(self, fname: str, body: str, pattern: str, sep: str) -> str | None:
        """
        Given a filename, body name, pattern string and separator,
        extract the {Subtype} token if the pattern matches.
        Returns the alias string, "" for no subtype, or None if pattern doesn't match.
        """
        import re as _re

        # Strip extension
        base = _re.sub(r"\.(pdf|mp4|mp3|docx?|srt|vtt)$", "", fname, flags=_re.IGNORECASE)

        # Split filename on separator
        parts = base.split(sep)

        # Parse pattern into tokens
        token_re = _re.compile(r"\{([^}]+)\}")
        tokens = token_re.findall(pattern)

        if not tokens:
            return None

        # We'll match tokens to parts greedily for multi-word tokens
        # Build a regex from the pattern
        # Each token type:
        #   {Date}         → matches date-like segments (digits + separators)
        #   {Meeting Body} → matches known body or body parts
        #   {File Type}    → matches known file type words
        #   {Subtype}      → captures whatever is here
        #   {Ignore}       → matches any single segment, discarded

        # Rejoin base and build capturing regex
        escaped_sep = _re.escape(sep)

        # Map each token to a regex pattern
        FILE_TYPES = r"(?:agenda|agendas|minutes|minute|video|videos|packet|packets|other|min|mins)"
        DATE_PAT   = (r"(?:\d{4}" + escaped_sep + r"\d{1,2}" + escaped_sep + r"\d{1,2}"
                      r"|\d{1,2}" + escaped_sep + r"\d{1,2}" + escaped_sep + r"\d{4}"
                      r"|\d{8}|\d{4})")

        seg = r"[^" + escaped_sep + r"]+"   # one segment (no sep)

        regex_parts = []
        group_names = []
        for i, tok in enumerate(tokens):
            tok_l = tok.strip().lower()
            if tok_l == "date":
                regex_parts.append(DATE_PAT)
            elif tok_l == "file type":
                regex_parts.append(FILE_TYPES)
            elif tok_l == "meeting body":
                # Match body as compact or spaced version
                body_variants = [
                    _re.escape(body),
                    _re.escape(body.replace(" ", sep)),
                    _re.escape(_re.sub(r"\s+", "", body)),
                ]
                regex_parts.append("(?:" + "|".join(body_variants) + ")")
            elif tok_l == "subtype":
                regex_parts.append(f"(?P<subtype>{seg})")
                group_names.append("subtype")
            elif tok_l == "ignore":
                regex_parts.append(seg)
            else:
                regex_parts.append(seg)

        full_regex = escaped_sep.join(regex_parts)
        # Allow extra segments before/after with sep
        full_regex = "(?:" + escaped_sep + ".*)?" .replace(".*", seg + "(?:" + escaped_sep + seg + ")*") + full_regex + "(?:" + escaped_sep + ".*)?"

        # Try simple direct match first (faster)
        m = _re.search(full_regex, base, _re.IGNORECASE)
        if not m:
            # Try matching just the token sequence anywhere in the split parts
            return self._fallback_token_match(parts, tokens, sep, body)

        try:
            return m.group("subtype") if "subtype" in m.groupdict() else ""
        except IndexError:
            return ""

    def _fallback_token_match(self, parts: list, tokens: list, sep: str, body: str) -> str | None:
        """
        Fallback: try to align pattern tokens to filename segments positionally.
        Returns alias string or None if can't match.
        """
        import re as _re

        body_compact = body.replace(" ", "").lower()
        body_sep     = body.replace(" ", sep).lower()

        # Merge body multi-word into single segment if needed
        # Try to find body in parts
        body_words = body.lower().split()
        merged = []
        i = 0
        while i < len(parts):
            # Try to match multi-word body starting at i
            matched_body = False
            for length in range(len(body_words), 0, -1):
                candidate = sep.join(parts[i:i+length]).lower()
                if candidate == body_sep or parts[i].lower() == body_compact:
                    merged.append(("body", sep.join(parts[i:i+length])))
                    i += length
                    matched_body = True
                    break
            if not matched_body:
                merged.append(("seg", parts[i]))
                i += 1

        FILE_NOISE = {"agenda","agendas","minutes","minute","video","videos",
                      "packet","packets","other","min","mins"}
        DATE_RE    = _re.compile(r"^\d{4}$|^\d{8}$|^\d{1,2}$")

        def classify(seg_val):
            sl = seg_val.lower()
            if sl in FILE_NOISE:   return "file type"
            if DATE_RE.match(sl):  return "date"
            return "seg"

        # Build classified merged list
        classified = []
        for kind, val in merged:
            if kind == "body":
                classified.append(("meeting body", val))
            else:
                classified.append((classify(val), val))

        # Now align tokens to classified parts
        tok_idx = 0
        cls_idx = 0
        result  = None

        while tok_idx < len(tokens) and cls_idx < len(classified):
            tok   = tokens[tok_idx].strip().lower()
            c_cls, c_val = classified[cls_idx]

            if tok == "ignore":
                tok_idx  += 1
                cls_idx  += 1
            elif tok == "subtype":
                if c_cls not in ("meeting body", "file type", "date"):
                    result   = c_val
                tok_idx  += 1
                cls_idx  += 1
            elif tok == c_cls:
                tok_idx  += 1
                cls_idx  += 1
            else:
                # Skip unmatched classified segment
                cls_idx  += 1

        return result if result is not None else ""

    def _scan_subtypes(self):
        """Scan Excel using the selected pattern to extract alias tokens."""
        import re
        path = self.input_file.get().strip()
        if not path or not Path(path).exists():
            messagebox.showwarning("No File", "Please select a valid input Excel file first.")
            return

        engine = self._load_engine()
        if not engine:
            return

        try:
            xl = pd.ExcelFile(path)
        except Exception as e:
            messagebox.showerror("Error", f"Could not read Excel:\n{e}")
            return

        tab_body_map = {tc.sheet_name.strip(): tc
                        for tc in self.tab_configs if tc.sheet_name.strip()}

        # found[(body, alias_token)] = True
        # alias_token is the raw leftover token after stripping body + date,
        # or "" for regular meetings with no extra token.
        found: Dict[tuple, bool] = {}

        FILE_COLS = {"agendas","agenda","minutes","minute","videos","video",
                     "packets","packet","agenda packet","agenda packets",
                     "other","others","other files","otherfiles"}

        for sname in xl.sheet_names:
            try:
                df = pd.read_excel(path, sheet_name=sname, dtype=str).fillna("")
            except Exception:
                continue
            forced = tab_body_map.get(sname)

            for col in df.columns:
                if col.strip().lower() not in FILE_COLS:
                    continue

                for raw in df[col].dropna():
                    raw = str(raw).strip()
                    if not raw:
                        continue

                    fname = _extract_filename(raw)

                    # Determine body first
                    if forced and forced.body_name.strip():
                        body = forced.body_name.strip()
                    else:
                        df_pref    = forced.date_format.strip() if forced and forced.date_format.strip() else "MM-DD-YYYY"
                        parsed_tmp = engine.parse_filename(fname, df_pref)
                        if parsed_tmp["is_error"] or not parsed_tmp.get("body"):
                            continue
                        body = parsed_tmp["body"]

                    # ── Extract alias using selected pattern ──
                    pattern, sep = self._get_active_pattern()
                    if pattern:
                        alias_token = self._extract_alias_from_pattern(fname, body, pattern, sep)
                        if alias_token is None:
                            continue   # pattern didn't match this file
                    else:
                        alias_token = ""

                    found[(body, alias_token)] = True

        if not found:
            messagebox.showinfo("No Data",
                                "No parseable files found.\n\n"
                                "Make sure Sheet Mapping is configured and "
                                "your Excel has Agendas / Minutes / Videos columns.")
            return

        # ── Collapse shared aliases into universal (blank body) rows ──
        # Count how many distinct bodies each alias token appears in
        alias_bodies: Dict[str, set] = {}
        for (body, alias_token) in found.keys():
            alias_bodies.setdefault(alias_token, set()).add(body)

        # Build collapsed rows: blank body if alias seen in >1 body, else keep body
        collapsed: Dict[tuple, str] = {}  # (resolved_body, alias_token) → resolved_body
        for (body, alias_token), _ in found.items():
            resolved_body = "" if len(alias_bodies[alias_token]) > 1 else body
            collapsed[(resolved_body, alias_token)] = resolved_body

        # De-duplicate against existing rows (keyed by body + alias)
        existing_keys = {(so.body.strip().lower(), so.alias.strip().lower())
                         for so in self.subtype_overrides}
        added = 0
        for (resolved_body, alias_token) in sorted(collapsed.keys()):
            if (resolved_body.lower(), alias_token.lower()) not in existing_keys:
                self.subtype_overrides.append(
                    SubtypeOverride(body=resolved_body, alias=alias_token, subtype=""))
                added += 1

        universal_count = sum(1 for (b, _) in collapsed if b == "")
        specific_count  = sum(1 for (b, _) in collapsed if b != "")

        # Auto-sort: universal (blank body) rows first, then body-specific;
        # within each group longest alias first. Only if not manually ordered.
        if not getattr(self, "_subtypes_manually_ordered", False):
            self.subtype_overrides.sort(
                key=lambda x: (" " + x.body.lower() if x.body else "",
                               -len(x.alias), x.alias.lower()))
        self._subtypes_manually_ordered = False  # reset after scan

        self._refresh_subtypes_display()
        self._switch_tab("Subtypes")
        messagebox.showinfo(
            "Scan Complete",
            f"Found {len(collapsed)} unique alias row(s) from {len(found)} body+alias combination(s).\n"
            f"Added {added} new row(s).\n\n"
            f"  {universal_count} universal row(s) — blank body, matches any body\n"
            f"  {specific_count} body-specific row(s) — only one body had this alias\n\n"
            "• Alias column (amber) = raw token found in filename\n"
            "• Subtype column (blue) = output name — fill this in\n"
            "• Blank body = universal (applies to all bodies)\n"
            "• Body-specific rows take priority over universal rows\n"
            "  Use the ▲▼ arrows to adjust priority manually.")

    def _add_subtype_row(self):
        self.subtype_overrides.append(SubtypeOverride())
        self._refresh_subtypes_display()

    def _remove_selected_subtypes(self):
        self.subtype_overrides = [so for so in self.subtype_overrides
                                  if not getattr(so, '_sel', False)]
        self._refresh_subtypes_display()

    def _clear_subtypes(self):
        if not self.subtype_overrides:
            return
        if messagebox.askyesno("Clear All", "Remove all subtype override rows?"):
            self.subtype_overrides.clear()
            self._refresh_subtypes_display()

    # ================================================================
    # TAB 4 — RUN
    # ================================================================
    def _build_run_tab(self):
        tab = self.tab_frames["Run"]

        ctk.CTkLabel(tab, text="Run Import",
                     font=ctk.CTkFont(size=20, weight="bold")
                     ).pack(anchor="w", padx=20, pady=(18,4))

        self.summary_card = ctk.CTkFrame(tab, fg_color=PANEL_BG, corner_radius=10)
        self.summary_card.pack(fill="x", padx=20, pady=8)
        self.summary_label = ctk.CTkLabel(
            self.summary_card,
            text="Configure Setup, Sheet Mapping, and Subtypes, then click Run.",
            font=ctk.CTkFont(size=13), text_color="gray", anchor="w", justify="left"
        )
        self.summary_label.pack(padx=16, pady=14, anchor="w")

        btn_row = ctk.CTkFrame(tab, fg_color="transparent")
        btn_row.pack(fill="x", padx=20, pady=12)

        self.dry_run_btn = ctk.CTkButton(
            btn_row, text="🔍  Dry Run", width=160, height=46,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color=GOLD, hover_color=GOLD_HOVER,
            command=self._run_dry_run
        )
        self.dry_run_btn.pack(side="left", padx=(0,12))

        self.run_btn = ctk.CTkButton(
            btn_row, text="▶  Run Import", width=180, height=46,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color=GREEN, hover_color=GREEN_HOVER,
            command=self._run_import
        )
        self.run_btn.pack(side="left", padx=(0,12))

        self.open_btn = ctk.CTkButton(
            btn_row, text="📂  Open Output", width=160, height=46,
            fg_color=GRAY, hover_color=GRAY_HOVER,
            state="disabled", command=self._open_output
        )
        self.open_btn.pack(side="left", padx=(0,12))

        self.open_errors_btn = ctk.CTkButton(
            btn_row, text="⚠  Open Errors", width=150, height=46,
            fg_color="#7B2D00", hover_color="#9a3800",
            state="disabled", command=self._open_errors
        )
        self.open_errors_btn.pack(side="left", padx=(0,12))

        self.rescan_btn = ctk.CTkButton(
            btn_row, text="↺  Re-scan", width=130, height=46,
            fg_color=TEAL, hover_color=TEAL_HOVER,
            command=self._rescan
        )
        self.rescan_btn.pack(side="left", padx=(0,12))

        ctk.CTkButton(
            btn_row, text="← Back to Subtypes", width=170, height=46,
            fg_color="#3a3a3a", hover_color="#4a4a4a",
            command=lambda: self._switch_tab("Subtypes")
        ).pack(side="right")

        self.progress = ctk.CTkProgressBar(tab, mode="indeterminate", height=12)
        self.progress.pack(fill="x", padx=20, pady=(0,8))
        self.progress.set(0)

        ctk.CTkLabel(tab, text="Log", font=ctk.CTkFont(size=13, weight="bold"),
                     anchor="w").pack(anchor="w", padx=20, pady=(8,2))
        self.log_box = ctk.CTkTextbox(tab, height=320,
                                       font=ctk.CTkFont(family="Courier", size=11))
        self.log_box.pack(fill="both", expand=True, padx=20, pady=(0,14))
        self.log_box.configure(state="disabled")

    # ================================================================
    # TAB 5 — RESULTS
    # ================================================================
    def _build_results_tab(self):
        tab = self.tab_frames["Results"]

        ctk.CTkLabel(tab, text="Results",
                     font=ctk.CTkFont(size=20, weight="bold")
                     ).pack(anchor="w", padx=20, pady=(18,4))

        self.stats_bar = ctk.CTkFrame(tab, fg_color=PANEL_BG, corner_radius=10)
        self.stats_bar.pack(fill="x", padx=20, pady=8)
        self.stat_rows      = ctk.CTkLabel(self.stats_bar, text="—  meeting rows",    text_color="gray")
        self.stat_errors    = ctk.CTkLabel(self.stats_bar, text="—  errors",          text_color="gray")
        self.stat_dupes     = ctk.CTkLabel(self.stats_bar, text="—  duplicates",      text_color="gray")
        self.stat_unmatched = ctk.CTkLabel(self.stats_bar, text="—  unmatched files", text_color="gray")
        for w in [self.stat_rows, self.stat_errors, self.stat_dupes, self.stat_unmatched]:
            w.pack(side="left", padx=20, pady=12)

        sub_bar = ctk.CTkFrame(tab, fg_color="transparent")
        sub_bar.pack(fill="x", padx=20, pady=(0,4))
        for label in ["Preview", "Error Log"]:
            ctk.CTkButton(sub_bar, text=label, width=130,
                          fg_color=BLUE, hover_color=BLUE_HOVER,
                          command=lambda l=label: self._switch_result_view(l)
                          ).pack(side="left", padx=4)

        self.preview_outer = ctk.CTkFrame(tab, fg_color=PANEL_BG, corner_radius=8)
        self.preview_outer.pack(fill="both", expand=True, padx=20, pady=(0,8))
        self._build_preview_tree()

        self.error_outer = ctk.CTkFrame(tab, fg_color=PANEL_BG, corner_radius=8)
        self._build_error_tree()

        btn_row = ctk.CTkFrame(tab, fg_color="transparent")
        btn_row.pack(fill="x", padx=20, pady=(0,12))
        ctk.CTkButton(btn_row, text="📂  Open Output File", width=180,
                      fg_color=GRAY, hover_color=GRAY_HOVER,
                      command=self._open_output).pack(side="left", padx=(0,10))
        ctk.CTkButton(btn_row, text="← Run Again", width=140,
                      fg_color="#3a3a3a", hover_color="#4a4a4a",
                      command=lambda: self._switch_tab("Run")).pack(side="right")

    def _build_preview_tree(self):
        import tkinter.ttk as ttk
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Custom.Treeview",
                        background="#2b2b2b", foreground="white",
                        fieldbackground="#2b2b2b", rowheight=24,
                        font=("Arial", 11))
        style.configure("Custom.Treeview.Heading",
                        background="#1f538d", foreground="white",
                        font=("Arial", 11, "bold"))
        style.map("Custom.Treeview", background=[("selected","#1f538d")])

        cols = ["Event Name","Event Date","Event Category","Meeting Type","Agenda","Minutes","Video","Packet"]
        self.preview_tree = ttk.Treeview(self.preview_outer, columns=cols,
                                          show="headings", style="Custom.Treeview")
        widths = [260, 90, 190, 120, 55, 55, 55, 55]
        for col, w in zip(cols, widths):
            self.preview_tree.heading(col, text=col)
            self.preview_tree.column(col, width=w, minwidth=40)
        vsb = ttk.Scrollbar(self.preview_outer, orient="vertical",
                            command=self.preview_tree.yview)
        self.preview_tree.configure(yscrollcommand=vsb.set)
        self.preview_tree.pack(side="left", fill="both", expand=True, padx=4, pady=4)
        vsb.pack(side="right", fill="y", pady=4)
        self.preview_tree.tag_configure("dup", background="#5a4a00")

    def _build_error_tree(self):
        import tkinter.ttk as ttk
        cols = ["Filename","Source","Issue"]
        self.error_tree = ttk.Treeview(self.error_outer, columns=cols,
                                        show="headings", style="Custom.Treeview")
        widths = [380, 90, 500]
        for col, w in zip(cols, widths):
            self.error_tree.heading(col, text=col)
            self.error_tree.column(col, width=w, minwidth=50)
        vsb2 = ttk.Scrollbar(self.error_outer, orient="vertical",
                             command=self.error_tree.yview)
        self.error_tree.configure(yscrollcommand=vsb2.set)
        self.error_tree.pack(side="left", fill="both", expand=True, padx=4, pady=4)
        vsb2.pack(side="right", fill="y", pady=4)

    def _switch_result_view(self, label: str):
        if label == "Preview":
            self.error_outer.pack_forget()
            self.preview_outer.pack(fill="both", expand=True, padx=20, pady=(0,8))
        else:
            self.preview_outer.pack_forget()
            self.error_outer.pack(fill="both", expand=True, padx=20, pady=(0,8))

    # ================================================================
    # FTP SCAN
    # ================================================================
    # Folder names that map to each file type (case-insensitive)
    _AGENDA_NAMES  = {"agendas", "agenda"}
    _MINUTES_NAMES = {"minutes", "minute", "min", "mins"}
    _VIDEO_NAMES   = {"videos", "video", "vid", "vids", "recordings", "recording"}
    _PACKET_NAMES  = {"packets", "packet", "agenda packet", "agenda packets",
                      "agendapacket", "agendapackets"}
    _OTHER_NAMES   = {"other", "others", "other files", "otherfiles",
                      "supporting", "supporting documents", "misc", "miscellaneous"}
    _CC_NAMES      = {"cc", "captions", "closed captions"}

    def _extract_dates_from_filename(self, fname: str) -> set:
        """
        Extract all date strings from a filename using every supported format.
        Returns a set of normalised date strings in YYYY-MM-DD form so that
        filenames using different separators / orderings still match each other.
        """
        import re
        from datetime import datetime as _dt

        base = re.sub(r"\.(srt|vtt|mp4|mp3|pdf|docx?)$", "", fname, flags=re.IGNORECASE)
        base = re.sub(r"[_\-\.]", " ", base)
        tokens = re.findall(r"\d+", base)

        dates = set()

        def _try(y, m, d):
            try:
                dates.add(_dt(int(y), int(m), int(d)).strftime("%Y-%m-%d"))
            except (ValueError, TypeError):
                pass

        # ── 8-digit runs: YYYYMMDD, MMDDYYYY, DDMMYYYY ──
        for tok in tokens:
            if len(tok) == 8:
                _try(tok[0:4], tok[4:6], tok[6:8])   # YYYYMMDD
                _try(tok[4:8], tok[0:2], tok[2:4])   # MMDDYYYY
                _try(tok[4:8], tok[2:4], tok[0:2])   # DDMMYYYY

        # ── 6-digit runs: YYMMDD, MMDDYY, DDMMYY ──
        for tok in tokens:
            if len(tok) == 6:
                try:
                    from Meeting_Import_Engine import pivot_year
                    _try(pivot_year(int(tok[0:2])), tok[2:4], tok[4:6])  # YYMMDD
                    _try(pivot_year(int(tok[4:6])), tok[0:2], tok[2:4])  # MMDDYY
                    _try(pivot_year(int(tok[4:6])), tok[2:4], tok[0:2])  # DDMMYY
                except Exception:
                    pass

        # ── Consecutive token triples (separated by spaces after normalisation) ──
        for i in range(len(tokens) - 2):
            a, b, c = tokens[i], tokens[i+1], tokens[i+2]
            la, lb, lc = len(a), len(b), len(c)
            # YYYY-MM-DD  /  YYYY-DD-MM
            if la == 4:
                _try(a, b, c)
                _try(a, c, b)
            # MM-DD-YYYY  /  DD-MM-YYYY
            if lc == 4:
                _try(c, a, b)
                _try(c, b, a)
            # YY-MM-DD  /  MM-DD-YY  /  DD-MM-YY
            if la == 2 and lb == 2 and lc == 2:
                try:
                    from Meeting_Import_Engine import pivot_year
                    _try(pivot_year(int(a)), b, c)
                    _try(pivot_year(int(c)), a, b)
                    _try(pivot_year(int(c)), b, a)
                except Exception:
                    pass

        return dates

    def _match_cc_file(self, video_fname: str, cc_files: list) -> list:
        """
        Given a video filename and a list of CC filenames (.srt / .vtt),
        return all CC filenames whose extracted dates overlap with the
        video filename's extracted dates.
        Returns [] if no match, [fname] for a clean match,
        or [fname1, fname2, ...] when multiple files match (conflict).
        """
        if not cc_files:
            return []
        video_dates = self._extract_dates_from_filename(video_fname)
        if not video_dates:
            return []
        matches = [
            cc for cc in cc_files
            if self._extract_dates_from_filename(cc) & video_dates
        ]
        return matches

    def _ftp_scan(self):
        """Connect to FTP, walk directory, build input Excel + pre-populate Sheet Mapping."""
        host     = self.ftp_host.get().strip()
        user     = self.ftp_user.get().strip()
        password = self.ftp_password.get().strip()
        ftp_root = self.ftp_path.get().strip() or "/"

        if not host:
            messagebox.showwarning("FTP", "Please enter a host / address.")
            return
        if not self.output_file.get().strip():
            messagebox.showwarning("FTP", "Please set an Output Excel path first.")
            return

        self.ftp_connect_btn.configure(state="disabled", text="Connecting...")
        self.ftp_status_label.configure(text="Connecting...", text_color="gray")

        flat = self.ftp_structure.get() == "No Folder Structure"
        pattern, sep = self._get_flat_pattern() if flat else (None, None)

        def _thread():
            try:
                if flat:
                    result = self._do_flat_ftp_scan(host, user, password, ftp_root,
                                                    pattern, sep)
                else:
                    result = self._do_ftp_scan(host, user, password, ftp_root)
                self.after(0, lambda r=result: self._on_ftp_done(r))
            except Exception as e:
                import traceback
                tb = traceback.format_exc()
                try:
                    self.after(0, lambda msg=str(e), t=tb: self._on_ftp_error(msg, t))
                except Exception:
                    print(f"[FTP ERROR] {e}\n{tb}")

        threading.Thread(target=_thread, daemon=True).start()

    def _list_dir(self, ftp, path, debug=None):
        """
        Return list of (name, is_dir) for the given FTP path.
        Tries MLSD first (modern, space-safe), falls back to LIST/NLST.
        Shared by both the structured scan (_do_ftp_scan) and the flat
        scan (_do_flat_ftp_scan) — do not re-nest this inside either one.

        `debug`, if provided, is a list that diagnostic lines get appended to.
        """
        def _log(msg):
            if debug is not None:
                debug.append(msg)

        items = []

        # ── Method 1: MLSD (RFC 3659) — fully space-safe ──
        try:
            for name, facts in ftp.mlsd(path):
                if name in (".", ".."):
                    continue
                is_dir = facts.get("type", "").lower() in ("dir", "cdir", "pdir")
                items.append((name, is_dir))
            _log(f"MLSD OK for '{path}': {len(items)} items")
            return items
        except Exception as e:
            _log(f"MLSD failed for '{path}': {e} — trying LIST")

        # ── Method 2: LIST with robust Unix/Windows parsing ──
        try:
            lines = []
            ftp.retrlines(f"LIST {path}", lines.append)
            for line in lines:
                line = line.strip()
                if not line:
                    continue

                is_dir = line[0].lower() == "d"

                # Windows-style: "MM-DD-YY  HH:MMAM <DIR>  Name With Spaces"
                if line[0].isdigit():
                    is_dir = "<DIR>" in line.upper()
                    # Name starts after "<DIR>" or file-size field
                    idx = line.upper().find("<DIR>")
                    if idx != -1:
                        name = line[idx + 5:].strip()
                    else:
                        # "MM-DD-YY  HH:MM  12345  filename"
                        parts = line.split(None, 3)
                        name = parts[3].strip() if len(parts) >= 4 else ""
                else:
                    # Unix-style: "drwxr-xr-x 2 user group 4096 Jan 15 10:30 Name With Spaces"
                    # Fields: perms links owner group size month day time/year name
                    # The name begins after the 8th whitespace-delimited token
                    parts = line.split(None, 8)
                    name = parts[8].strip() if len(parts) >= 9 else ""

                if name and name not in (".", ".."):
                    items.append((name, is_dir))

            _log(f"LIST OK for '{path}': {len(items)} items → {[n for n,_ in items[:5]]}")
            return items
        except Exception as e:
            _log(f"LIST failed for '{path}': {e} — trying NLST")

        # ── Method 3: NLST (names only, no is_dir info) ──
        try:
            names = ftp.nlst(path)
            for name in names:
                name = name.split("/")[-1]   # strip any path prefix the server includes
                if name not in (".", "..") and name:
                    items.append((name, False))
            _log(f"NLST OK for '{path}': {len(items)} items (is_dir unknown)")
            return items
        except Exception as e:
            _log(f"NLST failed for '{path}': {e}")

        return items

    def _do_ftp_scan(self, host: str, user: str, password: str, ftp_root: str) -> dict:
        """
        Walk FTP tree expecting structure:
            <ftp_root>/
                <Body Folder>/
                    <Agendas|Minutes|Videos folder>/
                        file1.pdf
        Returns dict with keys: bodies, unrecognized_folders, total_files, debug_log
        """
        import ftplib

        debug = []   # diagnostic lines returned to caller

        ftp = ftplib.FTP()
        ftp.connect(host, 21, timeout=30)
        welcome = ftp.login(user, password)
        debug.append(f"Login OK: {welcome.strip()}")
        ftp.set_pasv(True)

        # ── Try to cd to root so we know it exists ──
        try:
            ftp.cwd(ftp_root)
            pwd = ftp.pwd()
            debug.append(f"CWD to '{ftp_root}' OK  →  pwd={pwd}")
        except Exception as e:
            raise RuntimeError(f"Could not change to remote path '{ftp_root}': {e}")

        bodies       = []
        unrecognized = []
        total_files  = 0

        top_items = self._list_dir(ftp, ftp_root, debug)
        debug.append(f"Top-level items in '{ftp_root}': {[n for n,_ in top_items]}")

        for body_name, is_dir in top_items:
            if body_name.startswith("."):
                continue
            if not is_dir:
                debug.append(f"  Skipping non-dir at top level: {body_name}")
                continue

            body_path = ftp_root.rstrip("/") + "/" + body_name
            rec = {"body_name": body_name, "agendas": [], "minutes": [], "videos": [], "packets": [], "other": [], "cc": []}

            type_items = self._list_dir(ftp, body_path, debug)
            debug.append(f"  Body '{body_name}' has {len(type_items)} sub-items")

            for type_name, is_type_dir in type_items:
                if type_name.startswith("."):
                    continue
                type_path  = body_path + "/" + type_name
                key_lower  = type_name.strip().lower()

                if key_lower in self._AGENDA_NAMES:
                    col = "agendas"
                elif key_lower in self._MINUTES_NAMES:
                    col = "minutes"
                elif key_lower in self._VIDEO_NAMES:
                    col = "videos"
                elif key_lower in self._PACKET_NAMES:
                    col = "packets"
                elif key_lower in self._OTHER_NAMES:
                    col = "other"
                elif key_lower in self._CC_NAMES:
                    col = "cc"
                else:
                    unrecognized.append(f"{body_name}/{type_name}")
                    debug.append(f"    Unrecognized folder: {type_name}")
                    continue

                if is_type_dir:
                    file_items = self._list_dir(ftp, type_path, debug)
                    for fname, _ in file_items:
                        if not fname.startswith("."):
                            if col == "cc" and not fname.lower().endswith((".srt", ".vtt")):
                                debug.append(f"    CC folder: skipping non-caption file: {fname}")
                                continue
                            rec[col].append(fname)
                            total_files += 1
                else:
                    if col == "cc" and not type_name.lower().endswith((".srt", ".vtt")):
                        debug.append(f"    CC folder: skipping non-caption file: {type_name}")
                    else:
                        rec[col].append(type_name)
                        total_files += 1

            bodies.append(rec)

        try:
            ftp.quit()
        except Exception:
            pass

        return {
            "bodies":       bodies,
            "unrecognized": unrecognized,
            "total_files":  total_files,
            "debug_log":    debug,
        }

    def _do_flat_ftp_scan(self, host, user, password, ftp_root,
                            pattern, sep) -> dict:
        """FTP flat scan: list all files in root, parse body+type from filename using pattern."""
        import ftplib

        debug = []   # diagnostic lines returned to caller

        ftp = ftplib.FTP()
        ftp.connect(host, 21, timeout=20)
        welcome = ftp.login(user or "anonymous", password or "")
        debug.append(f"Login OK: {welcome.strip()}")
        ftp.set_pasv(True)

        # List all files (non-recursive at root level first, then one level deep)
        all_files = []
        def _collect(path):
            try:
                items = self._list_dir(ftp, path, debug)
                for name, is_dir in items:
                    if name.startswith("."):
                        continue
                    full = path.rstrip("/") + "/" + name
                    if is_dir:
                        _collect(full)
                    else:
                        all_files.append(name)
            except Exception as e:
                debug.append(f"Error listing '{path}': {e}")

        _collect(ftp_root)
        ftp.quit()
        debug.append(f"Collected {len(all_files)} file(s) under '{ftp_root}'")

        result = self._group_flat_files(all_files, pattern, sep, source="ftp",
                                        base_url=None, sas=None)
        # Merge the FTP-level log (connection/listing) with the pattern-matching
        # log from _group_flat_files so the diagnostic dialog shows the whole story.
        result["debug_log"] = debug + result.get("debug_log", [])
        return result

    def _do_flat_azure_scan(self, client_path, sas_token, pattern, sep,
                             account_host="civicclerk.blob.core.windows.net",
                             container="import") -> dict:
        """Azure flat scan: list all blobs under client_path, parse body+type from filename."""
        import urllib.request, urllib.parse
        import xml.etree.ElementTree as ET
        from urllib.parse import unquote

        BASE_URL = f"https://{account_host}/{container}"
        sas = sas_token.lstrip("?") if sas_token else ""
        prefix = client_path.strip("/")

        # List ALL blobs under prefix (no delimiter = flat recursive listing)
        list_params = urllib.parse.urlencode({
            "restype": "container",
            "comp":    "list",
            "prefix":  prefix + "/" if not prefix.endswith("/") else prefix,
        })
        url = BASE_URL + "?" + list_params + ("&" + sas if sas else "")

        with urllib.request.urlopen(url, timeout=30) as resp:
            xml_data = resp.read()

        root = ET.fromstring(xml_data)
        ns = {"a": root.tag.split("}")[0].lstrip("{")} if "}" in root.tag else {}

        all_files = []
        blob_urls = {}
        for blob in root.iter("{%s}Blob" % ns.get("a","") if ns else "Blob"):
            name_el = blob.find("{%s}Name" % ns.get("a","")) if ns else blob.find("Name")
            if name_el is not None:
                raw_path = name_el.text
                fname    = unquote(raw_path.split("/")[-1])
                if fname:
                    all_files.append(fname)
                    blob_url = f"{BASE_URL}/{urllib.parse.quote(raw_path)}"
                    if sas:
                        blob_url += "?" + sas
                    blob_urls[fname] = blob_url

        return self._group_flat_files(all_files, pattern, sep,
                                      source="azure", base_url=None,
                                      sas=None, blob_urls=blob_urls)

    def _group_flat_files(self, filenames, pattern, sep,
                          source="ftp", base_url=None, sas=None,
                          blob_urls=None) -> dict:
        """
        Given a flat list of filenames and a pattern, extract {Meeting Body}
        and {File Type} from each filename and group into bodies dict.
        """
        import re

        FILE_NOISE = {n for s in [self._AGENDA_NAMES, self._MINUTES_NAMES,
                                  self._VIDEO_NAMES, self._PACKET_NAMES,
                                  self._OTHER_NAMES] for n in s}

        TOKEN_RE = re.compile(r"\{([^}]+)\}", re.IGNORECASE)
        tokens   = TOKEN_RE.findall(pattern) if pattern else []

        bodies       = {}   # body_name → rec dict
        unrecognized = []
        total_files  = 0
        debug        = []

        for fname in filenames:
            # Skip non-media files
            if not re.search(r"\.(pdf|mp4|mp3|docx?|srt|vtt)$", fname, re.IGNORECASE):
                continue

            base = re.sub(r"\.(pdf|mp4|mp3|docx?|srt|vtt)$", "", fname, re.IGNORECASE)
            parts = base.split(sep)

            body_token = None
            type_token = None

            for i, tok in enumerate(tokens):
                tok_l = tok.strip().lower()
                if i >= len(parts):
                    break
                if tok_l == "meeting body":
                    body_token = parts[i]
                elif tok_l == "file type":
                    type_token = parts[i].lower()
                elif tok_l == "ignore":
                    continue
                # date and subtype tokens just consume their position

            if not body_token:
                unrecognized.append(fname)
                debug.append(f"  No body token: {fname}")
                continue

            # Normalise body token for matching
            body_norm  = body_token.replace("_"," ").replace("-"," ").strip()
            body_lower = body_norm.lower()

            # Determine file column from type token
            col = None
            if type_token:
                type_clean = type_token.replace("_"," ").strip()
                if type_clean in self._AGENDA_NAMES:    col = "agendas"
                elif type_clean in self._MINUTES_NAMES: col = "minutes"
                elif type_clean in self._VIDEO_NAMES:   col = "videos"
                elif type_clean in self._PACKET_NAMES:  col = "packets"
                elif type_clean in self._OTHER_NAMES:   col = "other"
            if col is None:
                # Fallback: guess from extension
                if fname.lower().endswith((".mp4",".mp3")):   col = "videos"
                elif fname.lower().endswith((".srt",".vtt")): col = "other"
                else:                                          col = "other"

            # Get or create body record
            if body_lower not in bodies:
                bodies[body_lower] = {"body_name": body_norm,
                                      "agendas":[], "minutes":[], "videos":[],
                                      "packets":[], "other":[]}

            file_ref = blob_urls.get(fname, fname) if blob_urls else fname
            bodies[body_lower][col].append(file_ref)
            total_files += 1

        return {
            "bodies":       list(bodies.values()),
            "unrecognized": unrecognized,
            "total_files":  total_files,
            "debug_log":    debug,
        }

    def _on_ftp_done(self, result: dict):
        self.ftp_connect_btn.configure(state="normal", text="🔌  Connect & Scan FTP")

        bodies      = result["bodies"]
        unrecog     = result["unrecognized"]
        total_files = result["total_files"]
        debug_log   = result.get("debug_log", [])

        if not bodies:
            diag = "\n".join(debug_log) if debug_log else "(no diagnostic info)"
            self.ftp_status_label.configure(
                text="⚠  No body folders found. Check the Remote Path.", text_color="#ffaaaa")
            messagebox.showwarning(
                "No Body Folders Found",
                "The FTP scan connected successfully but found no body folders.\n\n"
                "Common causes:\n"
                "• Remote Path is wrong — try setting it to the subfolder that contains\n"
                "  the meeting body folders (e.g. /NASSAUCOFL or /imports/NASSAUCOFL)\n"
                "• The server uses a different listing format\n\n"
                "── Diagnostic log ──\n" + diag
            )
            return

        # ── Build input Excel in memory ──
        # One sheet per body, columns: Agendas, Minutes, Videos
        output_path = self.output_file.get().strip()
        # Save alongside output as "<name>_FTP_Input.xlsx"
        p = Path(output_path)
        input_path = str(p.parent / (p.stem + "_FTP_Input.xlsx"))

        import openpyxl
        wb = openpyxl.Workbook()
        wb.remove(wb.active)   # remove default sheet

        for body in bodies:
            full_name  = body["body_name"]
            sheet_name = full_name[:31]   # Excel sheet name max 31 chars
            ws = wb.create_sheet(title=sheet_name)
            ws.append(["Agendas", "Minutes", "Videos", "Packets", "Other", "Close Captions"])
            cc_files = body.get("cc", [])
            max_rows = max(len(body["agendas"]), len(body["minutes"]),
                          len(body["videos"]), len(body.get("packets",[])),
                          len(body.get("other",[])), 1)
            for i in range(max_rows):
                video_fname = body["videos"][i] if i < len(body["videos"]) else ""
                cc_match    = ""
                if video_fname:
                    matched = self._match_cc_file(video_fname, cc_files)
                    if len(matched) == 1:
                        cc_match = matched[0]
                    elif len(matched) > 1:
                        cc_match = matched[0]
                        for extra in matched[1:]:
                            self._pending_errors.append({
                                "Filename": video_fname,
                                "Source":   "FTP CC",
                                "Issue":    f"Multiple CC files matched by date: '{extra}' conflicts with '{matched[0]}'",
                            })
                ws.append([
                    body["agendas"][i]              if i < len(body["agendas"])             else "",
                    body["minutes"][i]              if i < len(body["minutes"])             else "",
                    video_fname,
                    body.get("packets",[])[i]       if i < len(body.get("packets",[]))      else "",
                    body.get("other",[])[i]         if i < len(body.get("other",[]))        else "",
                    cc_match,
                ])

        wb.save(input_path)
        self.input_file.set(input_path)

        # ── Pre-populate Sheet Mapping ──
        existing_sheets = {tc.sheet_name for tc in self.tab_configs}
        added = 0
        for body in bodies:
            full_name  = body["body_name"]
            sheet_name = full_name[:31]
            if sheet_name not in existing_sheets:
                # sheet_name = truncated Excel name; body_name = full folder name
                self.tab_configs.append(TabConfig(
                    sheet_name=sheet_name,
                    body_name=full_name,
                ))
                added += 1

        self._refresh_mapping_display()

        # ── Status message ──
        msg_parts = [
            f"✓  Connected.  {len(bodies)} body folder(s) found, {total_files} files.",
            f"Input Excel saved: {Path(input_path).name}",
            f"Sheet Mapping pre-populated with {added} new row(s) — edit any field freely.",
        ]
        if unrecog:
            msg_parts.append(f"⚠  {len(unrecog)} unrecognized folder(s) skipped: {', '.join(unrecog[:5])}")

        self.ftp_status_label.configure(text=msg_parts[0], text_color="#aaffaa")

        detail = "\n".join(msg_parts)
        if unrecog:
            messagebox.showwarning("FTP Scan Complete", detail)
        else:
            messagebox.showinfo("FTP Scan Complete", detail)

        self._switch_tab("Sheet Mapping")


    def _build_flat_pattern_row(self, parent):
        """Build the pattern + separator + custom entry row for flat scan mode."""
        top_row = ctk.CTkFrame(parent, fg_color="transparent")
        top_row.pack(fill="x", pady=(0,4))
        ctk.CTkLabel(top_row, text="Filename Pattern:", anchor="e", width=130
                     ).pack(side="left", padx=(0,8))
        self._flat_pattern_menu = ctk.CTkOptionMenu(
            top_row, variable=self.flat_pattern, width=300,
            values=PATTERN_PRESETS,
            command=self._on_flat_pattern_select)
        self._flat_pattern_menu.pack(side="left", padx=(0,8))
        ctk.CTkLabel(top_row, text="Sep:", width=34, anchor="e"
                     ).pack(side="left", padx=(0,4))
        ctk.CTkOptionMenu(top_row, variable=self.flat_sep, width=60,
                          values=SEPARATOR_OPTIONS).pack(side="left", padx=(0,8))
        self._flat_custom_entry = ctk.CTkEntry(
            top_row, textvariable=self.flat_custom, width=220,
            placeholder_text="Custom pattern e.g. {Meeting Body}_{File Type}_{Date}")
        self._flat_custom_entry.pack(side="left")
        self._flat_custom_entry.configure(state="disabled")

        # Token key hint
        hint = ctk.CTkFrame(parent, fg_color="#1a2a3a", corner_radius=6)
        hint.pack(fill="x", pady=(4,0))
        ctk.CTkLabel(hint,
                     text="{Meeting Body}  →  the segment used to identify the meeting body\n"
                          "{File Type}     →  matches Agenda, Minutes, Video, Packet, Other\n"
                          "{Date}          →  matches date segments (YYYY, MM, DD or full date)\n"
                          "{Subtype}       →  optional subtype token (e.g. SpecialMeeting)\n"
                          "{Ignore}        →  matches any single segment, discards it (repeatable)",
                     anchor="w", justify="left",
                     text_color="#aaccff", font=ctk.CTkFont(size=11)
                     ).pack(padx=10, pady=6, anchor="w")

    def _on_flat_pattern_select(self, value):
        if value == "Custom Pattern...":
            self._flat_custom_entry.configure(state="normal")
        else:
            self._flat_custom_entry.configure(state="disabled")

    def _toggle_flat_pattern(self, value, flat_row):
        """Show/hide the flat pattern row based on structure selection."""
        if value == "No Folder Structure":
            flat_row.pack(fill="x", padx=14, pady=(4,4))
        else:
            flat_row.pack_forget()

    def _get_flat_pattern(self):
        """Return (pattern, separator) for flat scan."""
        pat = self.flat_pattern.get()
        sep = self.flat_sep.get()
        if pat == "Custom Pattern...":
            pat = self.flat_custom.get().strip()
        return pat, sep

    def _parse_az_url(self):
        """Parse the SAS URL and populate the Client Path field."""
        from urllib.parse import urlparse
        import re as _re
        raw = self.az_sas_url.get().strip()
        if not raw:
            return
        try:
            parsed = urlparse(raw)
            path_parts = parsed.path.lstrip("/").split("/")
            # Skip container (first segment)
            type_names = {n for s in [self._AGENDA_NAMES, self._MINUTES_NAMES,
                                      self._VIDEO_NAMES, self._PACKET_NAMES,
                                      self._OTHER_NAMES] for n in s}
            prefix_parts = []
            for part in path_parts[1:]:
                if not part:
                    continue
                if _re.search(r"\.(pdf|mp4|mp3|docx?|srt|vtt)$", part, _re.IGNORECASE):
                    break
                if part.lower() in type_names:
                    break
                prefix_parts.append(part)
            # Only take the FIRST segment as the client root —
            # deeper segments are likely body folder names
            if prefix_parts:
                self.az_parsed_path.set(prefix_parts[0])
        except Exception:
            pass

    def _azure_scan(self):
        """Connect to Azure Blob Storage via full SAS URL, walk Body structure, build input Excel."""
        from urllib.parse import urlparse, urlunparse

        raw_url = self.az_sas_url.get().strip()
        if not raw_url:
            messagebox.showwarning("Azure", "Please paste a SAS URL.")
            return
        if not self.output_file.get().strip():
            messagebox.showwarning("Azure", "Please set an Output Excel path first.")
            return

        # Parse the SAS URL for account, container, and token
        try:
            parsed_url   = urlparse(raw_url)
            account_host = parsed_url.hostname
            path_parts   = parsed_url.path.lstrip("/").split("/")
            container    = path_parts[0] if path_parts else "import"
            sas_token    = parsed_url.query
        except Exception as e:
            messagebox.showerror("Azure", f"Could not parse SAS URL:\n{e}")
            return

        # Auto-extract client path from URL:
        # Structure: /container/CLIENTCODE[/intermediate]/body/type/file
        # Strategy: take CLIENTCODE, then also include the next segment ONLY if
        # it looks like a known intermediate folder (files, media, import, uploads etc.)
        _INTERMEDIATE = {"files","file","media","uploads","upload","data","content","import"}
        import re as _re
        segs = [s for s in path_parts[1:] if s]
        flat = self.az_structure.get() == "No Folder Structure"
        if not segs:
            client_path = ""
        elif not flat and len(segs) >= 2 and segs[1].lower() in _INTERMEDIATE:
            client_path = f"{segs[0]}/{segs[1]}"   # Structured: LUISTEST/files
        else:
            client_path = segs[0]                   # Flat or no intermediate: LUISTEST
        if not client_path:
            messagebox.showwarning("Azure",
                                   "Could not determine client path from URL.\n"
                                   "Make sure the URL contains the customer folder after the container.")
            return

        self.az_connect_btn.configure(state="disabled", text="Connecting...")
        self.az_status_label.configure(text="Connecting...", text_color="gray")

        flat = self.az_structure.get() == "No Folder Structure"
        pattern, sep = self._get_flat_pattern() if flat else (None, None)

        def _thread():
            try:
                if flat:
                    result = self._do_flat_azure_scan(
                        client_path, sas_token, pattern, sep,
                        account_host=account_host, container=container)
                else:
                    result = self._do_azure_scan(client_path, sas_token,
                                                 account_host=account_host,
                                                 container=container)
                self.after(0, lambda r=result: self._on_azure_done(r))
            except Exception as e:
                import traceback
                tb  = traceback.format_exc()
                msg = str(e)
                self.after(0, lambda m=msg, t=tb: self._on_azure_error(m, t))

        threading.Thread(target=_thread, daemon=True).start()

    def _do_azure_scan(self, client_path: str, sas_token: str,
                          account_host: str = "civicclerk.blob.core.windows.net",
                          container: str = "import") -> dict:
        """
        Walk Azure Blob Storage expecting:
            <client_path>/<body_folder>/<type_folder>/<files>
        Uses Azure Blob REST list API (no SDK required).
        Returns same structure as _do_ftp_scan.
        """
        import urllib.request
        import urllib.parse
        import xml.etree.ElementTree as ET
        from urllib.parse import unquote

        BASE_URL  = f"https://{account_host}/{container}"

        # Normalise SAS token — strip leading ? if present
        sas = sas_token.lstrip("?") if sas_token else ""

        debug = []

        def _list_blobs(prefix: str, delimiter: str = "/") -> tuple:
            """
            List blobs under prefix with delimiter (virtual directory listing).
            Returns (subdirs, files) where subdirs are common prefixes (folders)
            and files are blob names.
            """
            # Build list params separately — NEVER merge SAS token params into
            # the query dict since Azure validates parameter order as part of the
            # HMAC signature. Append the SAS token as-is after the list params.
            list_params = urllib.parse.urlencode({
                "restype":   "container",
                "comp":      "list",
                "prefix":    prefix if prefix.endswith("/") else prefix + "/",
                "delimiter": delimiter,
            })
            if sas:
                url = BASE_URL + "?" + list_params + "&" + sas
            else:
                url = BASE_URL + "?" + list_params
            debug.append(f"LIST {prefix}/")
            try:
                with urllib.request.urlopen(url, timeout=20) as resp:
                    xml_data = resp.read()
            except Exception as e:
                debug.append(f"  → ERROR: {e}")
                raise

            root = ET.fromstring(xml_data)
            ns   = {"a": root.tag.split("}")[0].lstrip("{")} if "}" in root.tag else {}

            # Virtual subdirectories
            subdirs = []
            for cp in root.iter("{%s}BlobPrefix" % ns.get("a","") if ns else "BlobPrefix"):
                name_el = cp.find("{%s}Name" % ns.get("a","")) if ns else cp.find("Name")
                if name_el is not None:
                    # Strip trailing slash, then take last segment
                    raw = name_el.text.rstrip("/")
                    folder = unquote(raw.split("/")[-1])
                    subdirs.append(folder)

            # Actual blobs at this level
            files = []
            for blob in root.iter("{%s}Blob" % ns.get("a","") if ns else "Blob"):
                name_el = blob.find("{%s}Name" % ns.get("a","")) if ns else blob.find("Name")
                if name_el is not None:
                    fname = unquote(name_el.text.split("/")[-1])
                    if fname:
                        files.append(fname)

            debug.append(f"  → {len(subdirs)} subdirs, {len(files)} files")
            return subdirs, files

        # ── Walk top level: expect body folders ──
        # If a subfolder is not a type folder and contains no files but has its
        # own subfolders, treat it as a pass-through (e.g. "files") and go deeper.
        prefix = client_path.strip("/")
        top_names, _ = _list_blobs(prefix)

        if not top_names:
            debug.append("No subdirs at top level — no body folders found")
            return {"bodies": [], "unrecognized": [], "total_files": 0, "debug_log": debug}

        TYPE_NAMES = (self._AGENDA_NAMES | self._MINUTES_NAMES | self._VIDEO_NAMES
                      | self._PACKET_NAMES | self._OTHER_NAMES)

        # Check if top level is a pass-through folder (single non-type subfolder
        # with no files — e.g. "files" or "media")
        scan_prefix = prefix
        body_names  = top_names
        if len(top_names) == 1 and top_names[0].lower() not in TYPE_NAMES:
            deeper, _ = _list_blobs(f"{prefix}/{top_names[0]}")
            if deeper:
                scan_prefix = f"{prefix}/{top_names[0]}"
                body_names  = deeper
                debug.append(f"Pass-through folder detected: '{top_names[0]}' — scanning one level deeper")

        debug.append(f"Body level: {body_names}")

        bodies      = []
        unrecognized = []
        total_files  = 0

        for body_name in body_names:
            body_prefix = f"{scan_prefix}/{body_name}"
            type_folders, _ = _list_blobs(body_prefix)

            rec = {"body_name": body_name,
                   "agendas": [], "minutes": [], "videos": [], "packets": [], "other": [], "cc": []}

            for tf in type_folders:
                key = tf.strip().lower()
                if key in self._AGENDA_NAMES:
                    col = "agendas"
                elif key in self._MINUTES_NAMES:
                    col = "minutes"
                elif key in self._VIDEO_NAMES:
                    col = "videos"
                elif key in self._PACKET_NAMES:
                    col = "packets"
                elif key in self._OTHER_NAMES:
                    col = "other"
                elif key in self._CC_NAMES:
                    col = "cc"
                else:
                    unrecognized.append(f"{body_name}/{tf}")
                    continue

                type_prefix = f"{body_prefix}/{tf}"
                _, blob_files = _list_blobs(type_prefix)

                # Store just the filename in the input Excel
                for fname in blob_files:
                    if col == "cc" and not fname.lower().endswith((".srt", ".vtt")):
                        debug.append(f"    CC folder: skipping non-caption file: {fname}")
                        continue
                    rec[col].append(fname)
                    total_files += 1

            bodies.append(rec)
            debug.append(f"  Body '{body_name}': "
                         f"{len(rec['agendas'])}A {len(rec['minutes'])}M "
                         f"{len(rec['videos'])}V {len(rec['packets'])}P")

        return {
            "bodies":       bodies,
            "unrecognized": unrecognized,
            "total_files":  total_files,
            "debug_log":    debug,
        }

    def _on_azure_done(self, result: dict):
        self.az_connect_btn.configure(state="normal", text="☁  Connect & Scan Azure")

        bodies      = result["bodies"]
        unrecog     = result["unrecognized"]
        total_files = result["total_files"]
        debug_log   = result.get("debug_log", [])

        if not bodies:
            diag = "\n".join(debug_log) if debug_log else "(no diagnostic info)"
            self.az_status_label.configure(
                text="⚠  No body folders found. Check the Client Path.", text_color="#ffaaaa")
            messagebox.showwarning(
                "No Body Folders Found",
                "Azure scan connected but found no body folders.\n\n"
                "• Verify the Client Path points to the folder containing body subfolders\n"
                "• Check SAS token has list/read permissions\n\n"
                "── Diagnostic log ──\n" + diag)
            return

        # ── Build input Excel ──
        output_path = self.output_file.get().strip()
        p           = Path(output_path)
        input_path  = str(p.parent / (p.stem + "_Azure_Input.xlsx"))

        import openpyxl
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        for body in bodies:
            full_name  = body["body_name"]
            sheet_name = full_name[:31]   # Excel sheet name max 31 chars
            ws = wb.create_sheet(title=sheet_name)
            ws.append(["Agendas", "Minutes", "Videos", "Packets", "Other", "Close Captions"])
            cc_files = body.get("cc", [])
            max_rows = max(len(body["agendas"]), len(body["minutes"]),
                          len(body["videos"]), len(body.get("packets", [])),
                          len(body.get("other", [])), 1)
            for i in range(max_rows):
                video_fname = body["videos"][i] if i < len(body["videos"]) else ""
                cc_match    = ""
                if video_fname:
                    matched = self._match_cc_file(video_fname, cc_files)
                    if len(matched) == 1:
                        cc_match = matched[0]
                    elif len(matched) > 1:
                        cc_match = matched[0]
                        for extra in matched[1:]:
                            self._pending_errors.append({
                                "Filename": video_fname,
                                "Source":   "Azure CC",
                                "Issue":    f"Multiple CC files matched by date: '{extra}' conflicts with '{matched[0]}'",
                            })
                ws.append([
                    body["agendas"][i]           if i < len(body["agendas"])           else "",
                    body["minutes"][i]           if i < len(body["minutes"])           else "",
                    video_fname,
                    body.get("packets", [])[i]   if i < len(body.get("packets", []))   else "",
                    body.get("other", [])[i]     if i < len(body.get("other", []))     else "",
                    cc_match,
                ])

        wb.save(input_path)
        self.input_file.set(input_path)

        # ── Pre-populate Sheet Mapping ──
        existing_sheets = {tc.sheet_name for tc in self.tab_configs}
        added = 0
        for body in bodies:
            full_name  = body["body_name"]
            sheet_name = full_name[:31]
            if sheet_name not in existing_sheets:
                self.tab_configs.append(TabConfig(sheet_name=sheet_name, body_name=full_name))
                added += 1

        self._refresh_mapping_display()

        msg_parts = [
            f"✓  Connected.  {len(bodies)} body folder(s) found, {total_files} files.",
            f"Input Excel saved: {Path(input_path).name}",
            f"Sheet Mapping pre-populated with {added} new row(s).",
        ]
        if unrecog:
            msg_parts.append(f"⚠  {len(unrecog)} unrecognized folder(s) skipped: {', '.join(unrecog[:5])}")

        self.az_status_label.configure(text=msg_parts[0], text_color="#aaffaa")

        detail = "\n".join(msg_parts)
        if unrecog:
            messagebox.showwarning("Azure Scan Complete", detail)
        else:
            messagebox.showinfo("Azure Scan Complete", detail)

        self._switch_tab("Sheet Mapping")

    def _on_azure_error(self, msg: str, tb: str):
        self.az_connect_btn.configure(state="normal", text="☁  Connect & Scan Azure")
        self.az_status_label.configure(
            text=f"❌  Connection failed: {msg}", text_color="#ffaaaa")
        messagebox.showerror(
            "Azure Error",
            f"Could not connect or scan Azure Blob Storage:\n\n{msg}\n\n"
            "Check client path and SAS token.\n"
            "SAS token must have: Read + List permissions on the container.")

    def _on_ftp_error(self, msg: str, tb: str):
        self.ftp_connect_btn.configure(state="normal", text="🔌  Connect & Scan FTP")
        self.ftp_status_label.configure(
            text=f"❌  Connection failed: {msg}", text_color="#ffaaaa")
        messagebox.showerror("FTP Error",
                             f"Could not connect or scan FTP:\n\n{msg}\n\n"
                             "Check host, username, password, and remote path.")

    # ================================================================
    # SHEET SCANNING (from local Excel)
    # ================================================================
    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="Select Input Excel File",
            filetypes=[("Excel files","*.xlsx *.xls"),("All files","*.*")]
        )
        if path:
            self.input_file.set(path)
            p = Path(path)
            self.output_file.set(str(p.parent / "Meetings_Output.xlsx"))
            self._scan_sheets_silent()

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Save Output As", defaultextension=".xlsx",
            filetypes=[("Excel files","*.xlsx"),("All files","*.*")],
            initialfile="Meetings_Output.xlsx"
        )
        if path:
            self.output_file.set(path)
            self._try_auto_load_config(path)

    def _customer_config_path(self) -> Path | None:
        """Return the _Config.json path for the current output file, or None."""
        op = self.output_file.get().strip()
        if not op:
            return None
        p = Path(op)
        return p.parent / (p.stem + "_Config.json")

    def _auto_save_customer_config(self):
        """Silently save Sheet Mapping + Subtypes to _Config.json alongside output."""
        cfg_path = self._customer_config_path()
        if not cfg_path:
            return
        try:
            data = {
                "tab_configs":       [tc.to_dict() for tc in self.tab_configs],
                "subtype_overrides": [so.to_dict() for so in self.subtype_overrides],
                "date_fixes":        self.date_fixes,
                "dup_fixes":         {str(k): v for k, v in self.dup_fixes.items()},
                "body_fixes":        self.body_fixes,
            }
            cfg_path.write_text(json.dumps(data, indent=2))
        except Exception:
            pass

    def _try_auto_load_config(self, output_path: str):
        """Silently load _Config.json if it exists alongside the output file."""
        try:
            cfg_path = Path(output_path).parent / (Path(output_path).stem + "_Config.json")
            if cfg_path.exists():
                data = json.loads(cfg_path.read_text())
                self.tab_configs       = [TabConfig.from_dict(d)       for d in data.get("tab_configs", [])]
                self.subtype_overrides = [SubtypeOverride.from_dict(d) for d in data.get("subtype_overrides", [])]
                self.date_fixes        = data.get("date_fixes", {})
                self.body_fixes        = data.get("body_fixes", {})
                # dup_fixes keys are tuples serialised as strings — restore them
                raw_dup = data.get("dup_fixes", {})
                self.dup_fixes = {}
                for k_str, v in raw_dup.items():
                    try:
                        # Stored as "('file1', 'file2', 'field')"
                        import ast
                        self.dup_fixes[ast.literal_eval(k_str)] = v
                    except Exception:
                        pass
                self._refresh_mapping_display()
                self._refresh_subtypes_display()
        except Exception:
            pass

    def _load_customer_config(self):
        """Manually browse to a _Config.json and load it, replacing current settings."""
        path = filedialog.askopenfilename(
            title="Load Config JSON",
            filetypes=[("Config files","*_Config.json"),("JSON files","*.json"),("All files","*.*")]
        )
        if not path:
            return
        try:
            data = json.loads(Path(path).read_text())
            self.tab_configs       = [TabConfig.from_dict(d)       for d in data.get("tab_configs", [])]
            self.subtype_overrides = [SubtypeOverride.from_dict(d) for d in data.get("subtype_overrides", [])]
            self._refresh_mapping_display()
            self._refresh_subtypes_display()
            messagebox.showinfo("Config Loaded",
                                f"Loaded {len(self.tab_configs)} sheet mapping row(s) and "
                                f"{len(self.subtype_overrides)} subtype row(s) from:\n{Path(path).name}")
        except Exception as e:
            messagebox.showerror("Load Error", f"Could not load config:\n{e}")

    def _scan_output_folder(self):
        """Scan the output folder for existing Excel files to use as input."""
        op = self.output_file.get().strip()
        folder = str(Path(op).parent) if op else ""
        folder = filedialog.askdirectory(title="Select folder to scan", initialdir=folder or ".")
        if not folder:
            return
        patterns = ["*_FTP_Input.xlsx", "*_Azure_Input.xlsx", "*_Output.xlsx", "*_Import.xlsx", "*.xlsx"]
        found_files = []
        for pat in patterns:
            found_files.extend(Path(folder).glob(pat))
        # Remove duplicates, sort
        found_files = sorted(set(found_files), key=lambda p: p.name)
        if not found_files:
            messagebox.showinfo("Scan Folder", f"No Excel files found in:\n{folder}")
            return
        # Show selection dialog
        names = [f.name for f in found_files]
        sel = _PickDialog(self, "Select Input File", "Choose a file to use as input:", names)
        self.wait_window(sel)
        if sel.result is not None:
            chosen = found_files[sel.result]
            self.input_file.set(str(chosen))
            self._scan_sheets_silent()
            messagebox.showinfo("File Set", f"Input file set to:\n{chosen.name}")

    def _scan_sheets_silent(self):
        path = self.input_file.get().strip()
        if not path or not Path(path).exists():
            return
        try:
            xl = pd.ExcelFile(path)
            existing = {tc.sheet_name for tc in self.tab_configs}
            for name in xl.sheet_names:
                if name not in existing:
                    self.tab_configs.append(TabConfig(sheet_name=name))
            self._refresh_mapping_display()
        except Exception:
            pass

    def _scan_sheets(self):
        path = self.input_file.get().strip()
        if not path or not Path(path).exists():
            messagebox.showwarning("No File", "Please select a valid input Excel file first.")
            return
        try:
            xl = pd.ExcelFile(path)
        except Exception as e:
            messagebox.showerror("Error", f"Could not read Excel file:\n{e}")
            return
        existing = {tc.sheet_name for tc in self.tab_configs}
        added = 0
        for name in xl.sheet_names:
            if name not in existing:
                self.tab_configs.append(TabConfig(sheet_name=name))
                added += 1
        self._refresh_mapping_display()
        self._switch_tab("Sheet Mapping")
        if added:
            messagebox.showinfo("Sheets Scanned",
                                f"Found {len(xl.sheet_names)} sheet(s). Added {added} new row(s).\n\n"
                                "Fill in the Body Name column for each sheet.")
        else:
            messagebox.showinfo("Up to Date", "All sheets are already listed.")

    # ================================================================
    # ENGINE LOADER
    # ================================================================
    def _load_engine(self):
        engine_path = Path(__file__).parent / "Meeting_Import_Engine.py"
        if not engine_path.exists():
            messagebox.showerror("Engine Not Found",
                                 "Could not find Meeting_Import_Engine.py.\n"
                                 "Please place it in the same folder as this GUI.")
            return None
        import importlib.util
        spec   = importlib.util.spec_from_file_location("engine", engine_path)
        engine = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(engine)
        return engine

    # ================================================================
    # IMPORT
    # ================================================================
    def _validate(self) -> Optional[str]:
        if not self.input_file.get().strip():
            return "Please select an input Excel file."
        if not Path(self.input_file.get()).exists():
            return "Input file does not exist."
        if not self.output_file.get().strip():
            return "Please select an output file path."
        return None

    def _rescan(self):
        """Re-run FTP or Azure scan (whichever was last used) to rebuild the input Excel.
        Keeps all Sheet Mapping, Subtypes, and other settings intact."""
        mode = self._source_mode.get()
        if mode == "ftp":
            host     = self.ftp_host.get().strip()
            user     = self.ftp_user.get().strip()
            password = self.ftp_password.get().strip()
            ftp_root = self.ftp_path.get().strip() or "/"
            if not host:
                messagebox.showwarning("Re-scan", "No FTP host configured. Go to Setup and fill in FTP credentials.")
                return
            self.rescan_btn.configure(state="disabled", text="Scanning…")
            self._log("Re-scanning FTP...\n")
            def _thread():
                try:
                    result = self._do_ftp_scan(host, user, password, ftp_root)
                    self.after(0, lambda r=result: self._on_rescan_done(r, "ftp"))
                except Exception as e:
                    tb  = traceback.format_exc()
                    msg = str(e)
                    self.after(0, lambda m=msg, t=tb: self._on_rescan_error(m, t))
            threading.Thread(target=_thread, daemon=True).start()

        elif mode == "azure":
            az_path = self.az_path.get().strip().strip("/")
            az_sas  = self.az_sas.get().strip()
            if not az_path:
                messagebox.showwarning("Re-scan", "No Azure client path configured. Go to Setup and fill in Azure details.")
                return
            self.rescan_btn.configure(state="disabled", text="Scanning…")
            self._log("Re-scanning Azure...\n")
            def _thread():
                try:
                    result = self._do_azure_scan(az_path, az_sas)
                    self.after(0, lambda r=result: self._on_rescan_done(r, "azure"))
                except Exception as e:
                    tb  = traceback.format_exc()
                    msg = str(e)
                    self.after(0, lambda m=msg, t=tb: self._on_rescan_error(m, t))
            threading.Thread(target=_thread, daemon=True).start()

        else:
            messagebox.showinfo("Re-scan",
                                "Re-scan is only available when using FTP or Azure as the data source. Switch the source mode in Setup first.")

    def _on_rescan_done(self, result: dict, mode: str):
        self.rescan_btn.configure(state="normal", text="\u21ba  Re-scan")
        bodies      = result.get("bodies", [])
        total_files = result.get("total_files", 0)

        if not bodies:
            self._log("Re-scan found no body folders -- input Excel not updated.\n")
            messagebox.showwarning("Re-scan",
                                   "Re-scan connected but found no body folders.\n"
                                   "Check the remote path/client path in Setup.")
            return

        output_path = self.output_file.get().strip()
        p           = Path(output_path)
        suffix      = "_FTP_Input.xlsx" if mode == "ftp" else "_Azure_Input.xlsx"
        input_path  = str(p.parent / (p.stem + suffix))

        import openpyxl as _opxl
        wb = _opxl.Workbook()
        wb.remove(wb.active)
        for body in bodies:
            full_name  = body["body_name"]
            sheet_name = full_name[:31]
            ws = wb.create_sheet(title=sheet_name)
            ws.append(["Agendas", "Minutes", "Videos", "Packets", "Other", "Close Captions"])
            cc_files = body.get("cc", [])
            max_rows = max(len(body["agendas"]), len(body["minutes"]),
                          len(body["videos"]), len(body.get("packets",[])),
                          len(body.get("other",[])), 1)
            for i in range(max_rows):
                video_fname = body["videos"][i] if i < len(body["videos"]) else ""
                cc_match    = ""
                if video_fname:
                    matched = self._match_cc_file(video_fname, cc_files)
                    if len(matched) == 1:
                        cc_match = matched[0]
                    elif len(matched) > 1:
                        cc_match = matched[0]
                        for extra in matched[1:]:
                            self._pending_errors.append({
                                "Filename": video_fname,
                                "Source":   f"{mode.upper()} CC",
                                "Issue":    f"Multiple CC files matched by date: '{extra}' conflicts with '{matched[0]}'",
                            })
                ws.append([
                    body["agendas"][i]           if i < len(body["agendas"])           else "",
                    body["minutes"][i]           if i < len(body["minutes"])           else "",
                    video_fname,
                    body.get("packets",[])[i]    if i < len(body.get("packets",[]))    else "",
                    body.get("other",[])[i]      if i < len(body.get("other",[]))      else "",
                    cc_match,
                ])
        wb.save(input_path)
        self.input_file.set(input_path)

        self._log(
            f"Re-scan complete -- {len(bodies)} bodies, {total_files} files.\n"
            f"Input Excel updated: {Path(input_path).name}\n"
            "Sheet Mapping and Subtypes unchanged.\n")

    def _on_rescan_error(self, msg: str, tb: str):
        self.rescan_btn.configure(state="normal", text="\u21ba  Re-scan")
        self._log(f"Re-scan failed: {msg}\n")
        messagebox.showerror("Re-scan Error", f"Re-scan failed:\n\n{msg}")

    def _run_dry_run(self):
        err = self._validate()
        if err:
            messagebox.showwarning("Validation", err)
            return
        if self._running:
            return
        self._running = True
        self.dry_run_btn.configure(state="disabled", text="Running…")
        self.run_btn.configure(state="disabled")
        self.progress.start()
        self._log_clear()
        self._log("🔍  DRY RUN — no output file will be written.\n")
        threading.Thread(target=self._import_thread, kwargs={"dry_run": True}, daemon=True).start()

    def _run_import(self):
        err = self._validate()
        if err:
            messagebox.showwarning("Validation", err)
            return
        if self._running:
            return
        self._running = True
        self.run_btn.configure(state="disabled", text="Running…")
        self.dry_run_btn.configure(state="disabled")
        self.progress.start()
        self._log_clear()
        self._log("Starting import…\n")
        threading.Thread(target=self._import_thread, kwargs={"dry_run": False}, daemon=True).start()

    def _import_thread(self, dry_run: bool = False):
        try:
            result = self._execute_import(dry_run=dry_run)
            self.after(0, lambda r=result, d=dry_run: self._on_import_done(r, d))
        except Exception as e:
            import traceback
            tb  = traceback.format_exc()
            msg = str(e)   # capture now — Python clears 'e' after except block exits
            try:
                self.after(0, lambda m=msg, t=tb: self._on_import_error(m, t))
            except Exception:
                print(f"[IMPORT ERROR] {msg}\n{tb}")

    def _execute_import(self, dry_run: bool = False) -> dict:
        from openpyxl import load_workbook
        from openpyxl.styles import PatternFill, Font, Alignment
        from openpyxl.utils import get_column_letter

        input_path  = self.input_file.get().strip()
        output_path = self.output_file.get().strip()

        engine = self._load_engine()
        if not engine:
            raise RuntimeError("Engine not found.")

        tab_body_map: Dict[str, TabConfig] = {
            tc.sheet_name.strip(): tc
            for tc in self.tab_configs if tc.sheet_name.strip()
        }

        # body_name (lower) -> TabConfig, so a parser-detected body can still
        # inherit the folder's Event Category / Meeting Type / Time configured
        # in Meeting Bodies (universal/blank-body rule).
        body_name_map: Dict[str, TabConfig] = {
            tc.body_name.strip().lower(): tc
            for tc in self.tab_configs if tc.body_name.strip()
        }

        # Build alias lookup once.
        # Three tiers (checked in order):
        #   1. Body-specific rows with aliases  — most specific
        #   2. Universal rows with aliases      — blank body, matches any body
        #   3. Body-specific catch-all          — blank alias, specific body
        #   4. Universal catch-all              — blank body AND blank alias
        _body_specific  = []   # (body_lower, [aliases], so)  — body filled, alias filled
        _universal      = []   # (None,        [aliases], so)  — body blank,  alias filled
        _body_catchall  = {}   # body_lower → so               — body filled, alias blank
        _univ_catchall  = None # so                            — body blank,  alias blank

        for so in self.subtype_overrides:
            aliases = [a.strip().lower() for a in so.alias.split(",") if a.strip()]
            body_l  = so.body.strip().lower()
            has_body  = bool(so.body.strip())
            has_alias = bool(aliases)
            if has_body and has_alias:
                _body_specific.append((body_l, aliases, so))
            elif not has_body and has_alias:
                _universal.append((aliases, so))
            elif has_body and not has_alias:
                _body_catchall[body_l] = so
            else:
                _univ_catchall = so   # last blank+blank row wins

        # Sort each group longest-alias-first
        _body_specific.sort(key=lambda x: max(len(a) for a in x[1]), reverse=True)
        _universal.sort(key=lambda x: max(len(a) for a in x[0]), reverse=True)

        def _match_alias(body: str, fname: str):
            """Return matching SubtypeOverride for this body+filename, or None."""
            fname_lower = fname.lower()
            body_lower  = body.lower()
            # Tier 1 — body-specific alias match
            for row_body, aliases, so in _body_specific:
                if row_body != body_lower:
                    continue
                for alias in aliases:
                    if alias in fname_lower:
                        return so
            # Tier 2 — universal alias match (any body)
            for aliases, so in _universal:
                for alias in aliases:
                    if alias in fname_lower:
                        return so
            # Tier 3 — body-specific catch-all (blank alias)
            if body_lower in _body_catchall:
                return _body_catchall[body_lower]
            # Tier 4 — universal catch-all
            return _univ_catchall

        self._log(f"Input:  {input_path}")
        self._log(f"Output: {output_path}")
        self._log(f"Sheet mappings:   {len(tab_body_map)}")
        self._log(f"Subtype alias rows: {len(_body_specific) + len(_universal) + len(_body_catchall) + (1 if _univ_catchall else 0)}\n")

        xl          = pd.ExcelFile(input_path)
        sheet_names = xl.sheet_names
        self._log(f"Sheets: {sheet_names}\n")

        OUTPUT_COLUMNS = [
            "Event Name","Event Date","Event Time","Event Category",
            "Meeting Type","Event Description","Video File Name",
            "Agenda File Name","Minutes File Name","Agenda Packet File Name",
            "Close Caption File Name","Notice File Name","Other File Name",
            "External Media URL","Agenda is Published","Minutes is Published",
            "Agenda Packet is Published"
        ]

        all_meetings:    Dict[tuple, dict] = {}
        all_errors:      List[dict]        = []
        keys_with_dupes: set               = set()
        seen:            Dict[str, dict]   = {}

        def _process_sheet(sname, df_sheet, forced_body):
            col_map = {}
            for col in df_sheet.columns:
                cl = col.strip().lower()
                if cl in ("agendas","agenda"):                                      col_map["agenda"]  = col
                elif cl in ("minutes","minute"):                                    col_map["minutes"] = col
                elif cl in ("videos","video"):                                      col_map["video"]   = col
                elif cl in ("packets","packet","agenda packet","agenda packets",
                            "agendapacket","agendapackets"):                        col_map["packet"]  = col
                elif cl in ("other","others","other files","otherfiles"):           col_map["other"]   = col
                elif cl in ("close captions","close caption","closecaptions","closecaption","cc","captions"): col_map["cc"] = col

            def _process_col(col_key, file_field):
                nonlocal forced_body        # assigned on line below; declare before first read
                if col_key not in col_map:
                    return
                col_name = col_map[col_key]
                for raw in df_sheet[col_name].dropna():
                    raw = str(raw).strip()
                    if not raw:
                        continue

                    sid = f"{sname}::{col_name}"
                    seen.setdefault(sid, {})
                    if raw in seen[sid]:
                        all_errors.append({"Filename": raw,
                                           "Source Column": col_name,
                                           "Issue": f"Duplicate filename in {col_name}"})
                        continue
                    seen[sid][raw] = True

                    fname          = _extract_filename(raw)
                    date_fmt_pref  = (forced_body.date_format.strip()
                                      if forced_body and forced_body.date_format.strip()
                                      else "MM-DD-YYYY")

                    # ── Apply user date format fix if set ──
                    if fname in self.date_fixes:
                        fmt_str  = self.date_fixes[fname]
                        fixed_dt = _parse_with_format(fname, fmt_str)
                        if fixed_dt:
                            # Inject the fixed date into a synthetic parsed result
                            parsed = {"date": fixed_dt, "is_error": False, "error_reason": "",
                                      "body": None, "subtype": "", "match_key": None,
                                      "stored_value": fname}
                        else:
                            # Format still didn't work — parse normally
                            parsed = engine.parse_filename(fname, date_fmt_pref)
                    else:
                        parsed = engine.parse_filename(fname, date_fmt_pref)

                    # ── Apply user body fix if set ──
                    if fname in self.body_fixes:
                        fixed_sheet = self.body_fixes[fname]
                        forced_body = tab_body_map.get(fixed_sheet) or forced_body

                    # Determine body — forced wins, else engine, else skip
                    if forced_body and forced_body.body_name.strip():
                        body      = forced_body.body_name.strip()
                        base_mt   = forced_body.meeting_type.strip()
                        base_ec   = forced_body.event_category.strip()
                        base_time = forced_body.default_time
                        if parsed.get("date") is None:
                            all_errors.append({"Filename": fname,
                                               "Source Column": col_name,
                                               "Issue": parsed.get("error_reason") or
                                                        "Could not extract a date from filename"})
                            continue
                    else:
                        if parsed["is_error"]:
                            all_errors.append({"Filename": fname,
                                               "Source Column": col_name,
                                               "Issue": parsed["error_reason"]})
                            continue
                        body      = parsed.get("body") or "Unknown"
                        # Inherit folder settings from the matching Meeting Body
                        # (blank/universal rule) instead of leaving them empty.
                        _cfg      = body_name_map.get(body.strip().lower())
                        base_mt   = _cfg.meeting_type.strip()   if _cfg else ""
                        base_ec   = _cfg.event_category.strip() if _cfg else ""
                        base_time = _cfg.default_time           if _cfg else ""

                    # ── Alias-based subtype matching (replaces engine keyword list) ──
                    ovr = _match_alias(body, fname)
                    if ovr:
                        subtype    = ovr.subtype.strip()   # user-supplied display name
                        event_name = (f"{base_ec} {subtype}".strip() if subtype
                                      else f"{base_ec} Meeting")
                        ec         = ovr.override_event_category.strip() or base_ec
                        mt         = ovr.override_meeting_type.strip()   or base_mt
                        evt_time   = ovr.override_time.strip()           or base_time
                    else:
                        subtype    = ""
                        event_name = f"{base_ec} Meeting"
                        ec         = base_ec
                        mt         = base_mt
                        evt_time   = base_time

                    date_str = parsed["date"].strftime("%m/%d/%Y")
                    key = (body.lower(), date_str, subtype.lower(), sname)

                    if key not in all_meetings:
                        all_meetings[key] = {
                            "Event Name":    event_name,
                            "Event Date":    date_str,
                            "Event Time":    evt_time,
                            "Event Category": ec,
                            "Meeting Type":  mt,
                            "Event Description": "",
                            "Video File Name": "",
                            "Agenda File Name": "",
                            "Minutes File Name": "",
                            "Agenda Packet File Name": "",
                            "Close Caption File Name": "",
                            "Notice File Name": "",
                            "Other File Name": "",
                            "External Media URL": "",
                            "Agenda is Published": "",
                            "Minutes is Published": "",
                            "Agenda Packet is Published": "",
                            "_key": key,
                        }

                    rec = all_meetings[key]
                    if rec[file_field]:
                        existing = rec[file_field]
                        # Check for a user-defined duplicate fix
                        fix_key  = (existing, fname, file_field)
                        fix_key2 = (fname, existing, file_field)
                        fix = self.dup_fixes.get(fix_key) or self.dup_fixes.get(fix_key2)

                        if fix in ("other1",):
                            # Move existing file to Other File Name
                            rec["Other File Name"] = existing
                            rec[file_field]        = fname
                        elif fix in ("other2",):
                            # Move new (duplicate) file to Other File Name
                            rec["Other File Name"] = fname
                        elif fix == "keep2":
                            # Replace existing with new file
                            rec[file_field] = fname
                        else:
                            # Default: keep first, flag error
                            all_errors.append({
                                "Filename":      fname,
                                "Existing File": existing,
                                "Source Column": col_name,
                                "File Field":    file_field,
                                "Issue":         "Duplicate",
                                "Meeting Key":   str(key),
                            })
                            keys_with_dupes.add(key)
                    else:
                        rec[file_field] = fname

            _process_col("agenda",  "Agenda File Name")
            _process_col("minutes", "Minutes File Name")
            _process_col("video",   "Video File Name")
            _process_col("packet",  "Agenda Packet File Name")
            _process_col("other",   "Other File Name")
            _process_col("cc",      "Close Caption File Name")

        # Read all sheets at once using the open handle — avoids re-parsing the file per sheet
        all_dfs = pd.read_excel(xl, sheet_name=sheet_names, dtype=str)
        for sname in sheet_names:
            df_sheet = all_dfs[sname].fillna("")
            forced   = tab_body_map.get(sname)
            self._log(f"  Sheet '{sname}'" +
                      (f"  →  {forced.body_name}" if forced and forced.body_name else "  (auto-detect)"))
            _process_sheet(sname, df_sheet, forced)

        rows   = list(all_meetings.values())
        df_out = pd.DataFrame(rows, columns=OUTPUT_COLUMNS + ["_key"])
        df_out["_sort_date"] = pd.to_datetime(df_out["Event Date"], format="%m/%d/%Y", errors="coerce")
        df_out = df_out.sort_values(["Event Category","_sort_date"]).drop(columns=["_sort_date","_key"])

        df_errors = (pd.DataFrame(all_errors) if all_errors
                     else pd.DataFrame(columns=["Filename","Source Column","Issue"]))

        # ── Convert Event Time from internal HH:MM (24h) → HH:MM:SS AM/PM ──
        def _fmt_time(t):
            if not t or ":" not in str(t):
                return t
            try:
                from datetime import datetime as _dt
                parts = str(t).split(":")
                h, m = int(parts[0]), int(parts[1])
                return _dt(2000, 1, 1, h, m).strftime("%I:%M:%S %p").lstrip("0") or "12:00:00 AM"
            except Exception:
                return t
        df_out["Event Time"] = df_out["Event Time"].apply(_fmt_time)

        # Error log path: same folder as output, same stem + "_Errors.xlsx"
        output_p    = Path(output_path)
        errors_path = str(output_p.parent / (output_p.stem + "_Errors.xlsx"))

        self._log(f"\n✓  {len(df_out)} meeting rows")
        self._log(f"✓  {len(df_errors)} errors / flags")

        if dry_run:
            self._log("\n🔍  DRY RUN complete — no file written.")
        else:
            self._log("Writing output…")

            # ── Main import file (CivicClerk Import only) ──
            with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                df_out.to_excel(writer, sheet_name="CivicClerk Import", index=False)

            YELLOW    = PatternFill("solid", fgColor="FFFF00")
            HDR_FILL  = PatternFill("solid", fgColor="1F4E79")
            HDR_FONT  = Font(bold=True, color="FFFFFF", name="Arial", size=10)
            DATA_FONT = Font(name="Arial", size=10)

            wb = load_workbook(output_path)
            ws = wb["CivicClerk Import"]
            for cell in ws[1]:
                cell.fill = HDR_FILL; cell.font = HDR_FONT
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.row_dimensions[1].height = 30

            name_date_key = {(rec["Event Name"], rec["Event Date"]): k
                             for k, rec in all_meetings.items()}
            for row in ws.iter_rows(min_row=2):
                en = str(row[0].value) if row[0].value else ""
                ed = str(row[1].value) if row[1].value else ""
                mk = name_date_key.get((en, ed))
                for cell in row:
                    cell.font = DATA_FONT
                    cell.alignment = Alignment(vertical="center")
                    if mk in keys_with_dupes:
                        cell.fill = YELLOW
            for col in ws.columns:
                mx = max((len(str(c.value)) if c.value else 0) for c in col)
                ws.column_dimensions[get_column_letter(col[0].column)].width = min(mx+4, 60)
            ws.freeze_panes = "A2"
            wb.save(output_path)
            self._log(f"✓  Saved: {output_path}")

            # ── Separate error log file ──
            with pd.ExcelWriter(errors_path, engine="openpyxl") as writer:
                df_errors.to_excel(writer, sheet_name="Error Log", index=False)

            ERR_HDR_FILL = PatternFill("solid", fgColor="7B2D00")
            wb_err = load_workbook(errors_path)
            we     = wb_err["Error Log"]
            for cell in we[1]:
                cell.fill = ERR_HDR_FILL
                cell.font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            we.row_dimensions[1].height = 25
            for row in we.iter_rows(min_row=2):
                for cell in row:
                    cell.font = DATA_FONT
                    cell.alignment = Alignment(vertical="center")
            for col in we.columns:
                mx = max((len(str(c.value)) if c.value else 0) for c in col)
                we.column_dimensions[get_column_letter(col[0].column)].width = min(mx+4, 80)
            we.freeze_panes = "A2"
            wb_err.save(errors_path)
            self._log(f"✓  Errors saved: {errors_path}")

        return {
            "meeting_rows": len(df_out),
            "error_count":  len(df_errors),
            "dup_count":    len(keys_with_dupes),
            "unmatched":    len([r for r in df_errors.to_dict("records")
                                 if "Duplicate" not in r.get("Issue","")]),
            "records":      df_out.to_dict("records"),
            "errors":       df_errors.to_dict("records"),
            "dup_keys":     keys_with_dupes,
            "output_path":  output_path,
            "errors_path":  errors_path,
        }

    def _on_import_done(self, result: dict, dry_run: bool = False):
        self._running = False
        self.progress.stop(); self.progress.set(1)
        self.run_btn.configure(state="normal", text="▶  Run Import")
        self.dry_run_btn.configure(state="normal", text="🔍  Dry Run")
        if not dry_run:
            self.open_btn.configure(state="normal")
            self.open_errors_btn.configure(state="normal")
        self.last_result = result
        mode = "DRY RUN — no file written" if dry_run else "COMPLETE"
        self._log(f"\n{'='*55}\n  {mode}\n"
                  f"  {result['meeting_rows']:>6} meeting rows\n"
                  f"  {result['error_count']:>6} errors / flags\n"
                  f"  {result['dup_count']:>6} duplicate slots\n{'='*55}\n")
        if dry_run:
            self._log("Review Results tab, then click ▶ Run Import when ready.\n")
        prefix = "Dry run:  " if dry_run else "Last run:  "
        self.summary_label.configure(
            text=(f"{prefix}{result['meeting_rows']} rows  •  "
                  f"{result['error_count']} errors  •  "
                  f"{result['dup_count']} duplicate slots"),
            text_color="#ffdd88" if dry_run else "white")
        self._pending_errors = result.get("errors", [])
        self._populate_results(result)
        self._populate_errors_tab()
        self._switch_tab("Results")

    def _on_import_error(self, msg: str, tb: str):
        self._running = False
        self.progress.stop(); self.progress.set(0)
        self.run_btn.configure(state="normal", text="▶  Run Import")
        self.dry_run_btn.configure(state="normal", text="🔍  Dry Run")
        self._log(f"\n❌  ERROR: {msg}\n{tb}")
        messagebox.showerror("Import Error", f"An error occurred:\n\n{msg}")

    # ================================================================
    # RESULTS
    # ================================================================
    def _populate_results(self, result: dict):
        self.stat_rows.configure(
            text=f"  {result['meeting_rows']}  meeting rows  ", text_color="#aaffaa")
        self.stat_errors.configure(
            text=f"  {result['error_count']}  errors  ",
            text_color="#ffaaaa" if result['error_count'] else "#aaffaa")
        self.stat_dupes.configure(
            text=f"  {result['dup_count']}  dup slots  ",
            text_color="#ffddaa" if result['dup_count'] else "#aaffaa")
        self.stat_unmatched.configure(
            text=f"  {result['unmatched']}  unmatched  ",
            text_color="#ffddaa" if result['unmatched'] else "#aaffaa")

        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        dup_keys = result.get("dup_keys", set())

        def _cell(v):
            # Render missing values (None / pandas NaN) as blank, never "nan".
            if v is None:
                return ""
            s = str(v)
            return "" if s.strip().lower() == "nan" else s

        for rec in result.get("records", []):
            tag = "dup" if rec.get("_key") in dup_keys else ""
            self.preview_tree.insert("", "end", tags=(tag,), values=(
                _cell(rec.get("Event Name","")),
                _cell(rec.get("Event Date","")),
                _cell(rec.get("Event Category","")),
                _cell(rec.get("Meeting Type","")),
                "✓" if rec.get("Agenda File Name") else "—",
                "✓" if rec.get("Minutes File Name") else "—",
                "✓" if rec.get("Video File Name") else "—",
                "✓" if rec.get("Agenda Packet File Name") else "—",
            ))

        for item in self.error_tree.get_children():
            self.error_tree.delete(item)
        for err in result.get("errors", []):
            self.error_tree.insert("", "end", values=(
                err.get("Filename",""),
                err.get("Source Column",""),
                err.get("Issue",""),
            ))

    # ================================================================
    # OPEN OUTPUT
    # ================================================================
    def _open_output(self):
        path = self.output_file.get().strip()
        if path and Path(path).exists():
            import platform
            if platform.system() == "Windows":
                os.startfile(path)
            elif platform.system() == "Darwin":
                sp.run(["open", path])
            else:
                sp.run(["xdg-open", path])
        else:
            messagebox.showwarning("No Output", "No output file found. Run import first.")

    def _open_errors(self):
        if not self.last_result:
            messagebox.showwarning("No Data", "Run import first.")
            return
        path = self.last_result.get("errors_path", "")
        if path and Path(path).exists():
            import platform
            if platform.system() == "Windows":
                os.startfile(path)
            elif platform.system() == "Darwin":
                sp.run(["open", path])
            else:
                sp.run(["xdg-open", path])
        else:
            messagebox.showwarning("No Errors File", "Errors file not found. Run import first.")

    # ================================================================
    # LOG
    # ================================================================
    # ================================================================
    # ERROR TAB
    # ================================================================
    def _build_errors_tab(self):
        from tkinter import ttk
        tab = self.tab_frames["Errors"]

        hdr = ctk.CTkFrame(tab, fg_color="transparent")
        hdr.pack(fill="x", padx=20, pady=(16,4))
        ctk.CTkLabel(hdr, text="Error Resolution",
                     font=ctk.CTkFont(size=20, weight="bold")).pack(side="left")
        ctk.CTkButton(hdr, text="Clear Fixes", width=120,
                      fg_color=GRAY, hover_color=GRAY_HOVER,
                      command=self._clear_all_fixes).pack(side="right")

        ctk.CTkLabel(tab,
                     text="Fix errors below, then click Run Import again. "
                          "Resolved errors disappear after the next run.",
                     text_color="gray", font=ctk.CTkFont(size=11)
                     ).pack(anchor="w", padx=20, pady=(0,8))

        # ── Date failures section ──
        date_hdr = ctk.CTkFrame(tab, fg_color=PANEL_BG, corner_radius=8)
        date_hdr.pack(fill="x", padx=20, pady=(0,2))
        ctk.CTkLabel(date_hdr, text="Date Parse Failures",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color="#ffaaaa").pack(side="left", padx=12, pady=6)

        date_col_hdr = ctk.CTkFrame(tab, fg_color="#2a1a1a", height=28)
        date_col_hdr.pack(fill="x", padx=20)
        date_col_hdr.pack_propagate(False)
        date_col_hdr.grid_columnconfigure(0, weight=3, minsize=300)
        date_col_hdr.grid_columnconfigure(1, weight=1, minsize=160)
        date_col_hdr.grid_columnconfigure(2, weight=2, minsize=200)
        date_col_hdr.grid_columnconfigure(3, minsize=180)
        for c, t in enumerate(["Filename", "Source Column", "Issue", "Date Format Override"]):
            ctk.CTkLabel(date_col_hdr, text=t, font=ctk.CTkFont(size=11, weight="bold"),
                         anchor="w").grid(row=0, column=c, padx=6, pady=4, sticky="ew")

        self.date_errors_scroll = ctk.CTkScrollableFrame(tab, height=160)
        self.date_errors_scroll.pack(fill="x", padx=20, pady=(0,10))

        # ── Duplicates section ──
        dup_hdr = ctk.CTkFrame(tab, fg_color=PANEL_BG, corner_radius=8)
        dup_hdr.pack(fill="x", padx=20, pady=(0,2))
        ctk.CTkLabel(dup_hdr, text="Duplicate File Conflicts",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color="#ffddaa").pack(side="left", padx=12, pady=6)

        dup_col_hdr = ctk.CTkFrame(tab, fg_color="#2a2200", height=28)
        dup_col_hdr.pack(fill="x", padx=20)
        dup_col_hdr.pack_propagate(False)
        dup_col_hdr.grid_columnconfigure(0, weight=2, minsize=200)
        dup_col_hdr.grid_columnconfigure(1, weight=2, minsize=200)
        dup_col_hdr.grid_columnconfigure(2, weight=1, minsize=130)
        dup_col_hdr.grid_columnconfigure(3, minsize=220)
        for c, t in enumerate(["File 1 (kept)", "File 2 (duplicate)", "File Type", "Resolution"]):
            ctk.CTkLabel(dup_col_hdr, text=t, font=ctk.CTkFont(size=11, weight="bold"),
                         anchor="w").grid(row=0, column=c, padx=6, pady=4, sticky="ew")

        self.dup_errors_scroll = ctk.CTkScrollableFrame(tab, height=160)
        self.dup_errors_scroll.pack(fill="x", padx=20, pady=(0,10))

        # ── Unmatched body section ──
        body_hdr = ctk.CTkFrame(tab, fg_color=PANEL_BG, corner_radius=8)
        body_hdr.pack(fill="x", padx=20, pady=(0,2))
        ctk.CTkLabel(body_hdr, text="Unmatched Body / Unrecognized Files",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color="#aaccff").pack(side="left", padx=12, pady=6)

        body_col_hdr = ctk.CTkFrame(tab, fg_color="#1a1a2a", height=28)
        body_col_hdr.pack(fill="x", padx=20)
        body_col_hdr.pack_propagate(False)
        body_col_hdr.grid_columnconfigure(0, weight=3, minsize=300)
        body_col_hdr.grid_columnconfigure(1, weight=1, minsize=160)
        body_col_hdr.grid_columnconfigure(2, minsize=220)
        for c, t in enumerate(["Filename", "Issue", "Assign to Body (Sheet Name)"]):
            ctk.CTkLabel(body_col_hdr, text=t,
                         font=ctk.CTkFont(size=11, weight="bold"),
                         anchor="w").grid(row=0, column=c, padx=6, pady=4, sticky="ew")

        # Pack btn_row BEFORE the expanding scroll frame so it's never pushed off screen
        btn_row = ctk.CTkFrame(tab, fg_color="transparent")
        btn_row.pack(fill="x", padx=20, pady=8)
        ctk.CTkButton(btn_row, text="← Back to Run", width=140,
                      fg_color="#3a3a3a", hover_color="#4a4a4a",
                      command=lambda: self._switch_tab("Run")
                      ).pack(side="left", padx=(0,12))
        ctk.CTkButton(btn_row, text="🔍  Re-run Dry Run", width=180,
                      fg_color=GOLD, hover_color=GOLD_HOVER,
                      font=ctk.CTkFont(size=13, weight="bold"),
                      command=self._errors_tab_dry_run
                      ).pack(side="left")

        self.body_errors_scroll = ctk.CTkScrollableFrame(tab, height=160)
        self.body_errors_scroll.pack(fill="both", expand=True, padx=20, pady=(0,10))

    def _errors_tab_dry_run(self):
        """Trigger a dry run from the Errors tab and stay on Errors tab after."""
        err = self._validate()
        if err:
            messagebox.showwarning("Validation", err)
            return
        if self._running:
            return
        self._running = True
        self.progress.start()
        self._log_clear()
        self._log("Re-running dry run from Errors tab...\n")
        def _done(result, dry_run):
            self._on_import_done(result, dry_run)
            self._switch_tab("Errors")   # stay on Errors tab
        def _thread():
            try:
                result = self._execute_import(dry_run=True)
                self.after(0, lambda r=result: _done(r, True))
            except Exception as e:
                import traceback
                tb  = traceback.format_exc()
                msg = str(e)
                self.after(0, lambda m=msg, t=tb: self._on_import_error(m, t))
        threading.Thread(target=_thread, daemon=True).start()

    def _populate_errors_tab(self):
        """Populate Error tab from self._pending_errors, skipping already-fixed entries."""
        # Clear existing rows
        for w in self.date_errors_scroll.winfo_children():
            w.destroy()
        for w in self.dup_errors_scroll.winfo_children():
            w.destroy()
        self._populate_body_errors()

        date_count = 0
        dup_count  = 0

        for err in self._pending_errors:
            issue = err.get("Issue", "")
            fname = err.get("Filename", "")

            if issue == "Duplicate":
                existing   = err.get("Existing File", "")
                file_field = err.get("File Field", "")
                fix_key    = (existing, fname, file_field)

                # Skip if already fixed
                if fix_key in self.dup_fixes or (fname, existing, file_field) in self.dup_fixes:
                    continue

                dup_count += 1
                color = ROW_A if dup_count % 2 == 0 else ROW_B
                row = ctk.CTkFrame(self.dup_errors_scroll, fg_color=color)
                row.pack(fill="x", pady=2)
                row.grid_columnconfigure(0, weight=2, minsize=200)
                row.grid_columnconfigure(1, weight=2, minsize=200)
                row.grid_columnconfigure(2, weight=1, minsize=130)
                row.grid_columnconfigure(3, minsize=220)

                ctk.CTkLabel(row, text=existing, anchor="w", font=ctk.CTkFont(size=11)
                             ).grid(row=0, column=0, padx=6, pady=5, sticky="ew")
                ctk.CTkLabel(row, text=fname, anchor="w", font=ctk.CTkFont(size=11),
                             text_color="#ffddaa"
                             ).grid(row=0, column=1, padx=6, pady=5, sticky="ew")
                ctk.CTkLabel(row, text=file_field.replace(" File Name",""), anchor="w",
                             font=ctk.CTkFont(size=11)
                             ).grid(row=0, column=2, padx=6, pady=5, sticky="ew")

                res_var = ctk.StringVar(value="Keep File 1")
                ctk.CTkOptionMenu(row, variable=res_var, width=215,
                                  values=["Keep File 1",
                                          "Keep File 2",
                                          "File 1 → Other File Name",
                                          "File 2 → Other File Name"],
                                  command=lambda v, fk=fix_key, rv=res_var:
                                      self._apply_dup_fix(fk, rv.get())
                                  ).grid(row=0, column=3, padx=6, pady=5, sticky="w")

            else:
                # Date / other parse error — skip if already fixed
                if fname in self.date_fixes:
                    continue

                date_count += 1
                color = ROW_A if date_count % 2 == 0 else ROW_B
                row = ctk.CTkFrame(self.date_errors_scroll, fg_color=color)
                row.pack(fill="x", pady=2)
                row.grid_columnconfigure(0, weight=3, minsize=300)
                row.grid_columnconfigure(1, weight=1, minsize=160)
                row.grid_columnconfigure(2, weight=2, minsize=200)
                row.grid_columnconfigure(3, minsize=180)

                ctk.CTkLabel(row, text=fname, anchor="w", font=ctk.CTkFont(size=11)
                             ).grid(row=0, column=0, padx=6, pady=5, sticky="ew")
                ctk.CTkLabel(row, text=err.get("Source Column",""), anchor="w",
                             font=ctk.CTkFont(size=11)
                             ).grid(row=0, column=1, padx=6, pady=5, sticky="ew")
                ctk.CTkLabel(row, text=issue, anchor="w", font=ctk.CTkFont(size=11),
                             text_color="#ffaaaa"
                             ).grid(row=0, column=2, padx=6, pady=5, sticky="ew")

                DATE_FORMATS = [
                    "(select format)",
                    "MM-DD-YYYY  (e.g. 01-15-2025)",
                    "DD-MM-YYYY  (e.g. 15-01-2025)",
                    "YYYY-MM-DD  (e.g. 2025-01-15)",
                    "MM-DD-YY    (e.g. 01-15-25)",
                    "DD-MM-YY    (e.g. 15-01-25)",
                    "YY-MM-DD    (e.g. 25-01-15)",
                    "MMDDYYYY    (e.g. 01152025)",
                    "DDMMYYYY    (e.g. 15012025)",
                    "YYYYMMDD    (e.g. 20250115)",
                    "YYMMDD      (e.g. 250115)",
                ]
                fmt_var = ctk.StringVar(value=self.date_fixes.get(fname, "(select format)"))
                ctk.CTkOptionMenu(row, variable=fmt_var, width=210,
                                  values=DATE_FORMATS,
                                  command=lambda v, fn=fname: self._apply_date_fix(fn, v)
                                  ).grid(row=0, column=3, padx=6, pady=5, sticky="w")

    def _populate_body_errors(self):
        """Populate unmatched body errors section."""
        if not hasattr(self, "body_errors_scroll"):
            return
        for w in self.body_errors_scroll.winfo_children():
            w.destroy()

        body_count = 0
        sheet_names = ["(select body)"] + [
            tc.sheet_name for tc in self.tab_configs if tc.sheet_name.strip()]

        for err in self._pending_errors:
            issue = err.get("Issue", "")
            fname = err.get("Filename", "")
            if "Unrecognized" not in issue and "body" not in issue.lower() and "pattern" not in issue.lower():
                continue
            if fname in self.body_fixes:
                continue

            body_count += 1
            color = ROW_A if body_count % 2 == 0 else ROW_B
            row = ctk.CTkFrame(self.body_errors_scroll, fg_color=color)
            row.pack(fill="x", pady=2)
            row.grid_columnconfigure(0, weight=3, minsize=300)
            row.grid_columnconfigure(1, weight=1, minsize=160)
            row.grid_columnconfigure(2, minsize=220)

            ctk.CTkLabel(row, text=fname, anchor="w",
                         font=ctk.CTkFont(size=11)
                         ).grid(row=0, column=0, padx=6, pady=5, sticky="ew")
            ctk.CTkLabel(row, text=issue, anchor="w",
                         font=ctk.CTkFont(size=11), text_color="#aaccff"
                         ).grid(row=0, column=1, padx=6, pady=5, sticky="ew")

            body_var = ctk.StringVar(value="(select body)")
            ctk.CTkOptionMenu(row, variable=body_var, width=215,
                              values=sheet_names,
                              command=lambda v, fn=fname: self._apply_body_fix(fn, v)
                              ).grid(row=0, column=2, padx=6, pady=5, sticky="w")

    def _apply_body_fix(self, fname: str, sheet_name: str):
        """Store body fix for a filename."""
        if sheet_name and sheet_name != "(select body)":
            self.body_fixes[fname] = sheet_name
            self._auto_save_customer_config()
            self._populate_body_errors()

    def _apply_date_fix(self, fname: str, fmt: str):
        """Store date format fix for a filename."""
        if fmt and fmt != "(select format)":
            self.date_fixes[fname] = fmt
            self._auto_save_customer_config()
        else:
            self.date_fixes.pop(fname, None)

    def _apply_dup_fix(self, fix_key: tuple, resolution: str):
        """Store duplicate fix."""
        mapping = {
            "Keep File 1":              "keep1",
            "Keep File 2":              "keep2",
            "File 1 → Other File Name": "other1",
            "File 2 → Other File Name": "other2",
        }
        code = mapping.get(resolution, "keep1")
        if code == "keep1":
            self.dup_fixes.pop(fix_key, None)
        else:
            self.dup_fixes[fix_key] = code
        self._auto_save_customer_config()

    def _clear_all_fixes(self):
        if messagebox.askyesno("Clear Fixes", "Remove all date and duplicate fixes?"):
            self.date_fixes.clear()
            self.dup_fixes.clear()
            self.body_fixes.clear()
            self._auto_save_customer_config()
            self._populate_errors_tab()

    def _log(self, text: str):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _log_clear(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0","end")
        self.log_box.configure(state="disabled")

    # ================================================================
    # CONFIG PERSISTENCE
    # ================================================================
    def _save_config(self):
        try:
            cfg = {
                "input_file":        self.input_file.get(),
                "output_file":       self.output_file.get(),
                "tab_configs":       [tc.to_dict() for tc in self.tab_configs],
                "subtype_overrides": [so.to_dict() for so in self.subtype_overrides],
                "ftp_host":          self.ftp_host.get(),
                "ftp_user":          self.ftp_user.get(),
                "ftp_path":          self.ftp_path.get(),
                "az_sas_url":        self.az_sas_url.get(),
                # Note: passwords/SAS tokens intentionally NOT saved for security
            }
            CONFIG_FILE.write_text(json.dumps(cfg, indent=2))
        except Exception:
            pass

    def _load_config(self):
        if not CONFIG_FILE.exists():
            return
        try:
            cfg = json.loads(CONFIG_FILE.read_text())
            self.input_file.set(cfg.get("input_file",""))
            self.output_file.set(cfg.get("output_file",""))
            self.tab_configs       = [TabConfig.from_dict(d)       for d in cfg.get("tab_configs",[])]
            self.subtype_overrides = [SubtypeOverride.from_dict(d) for d in cfg.get("subtype_overrides",[])]
            self.ftp_host.set(cfg.get("ftp_host",""))
            self.ftp_user.set(cfg.get("ftp_user",""))
            self.ftp_path.set(cfg.get("ftp_path","/"))
            self.az_sas_url.set(cfg.get("az_sas_url",""))
            self._refresh_mapping_display()
            self._refresh_subtypes_display()
        except Exception:
            pass

    def destroy(self):
        self._save_config()
        super().destroy()


class _PickDialog(ctk.CTkToplevel):
    """Simple modal list-picker dialog."""
    def __init__(self, parent, title, prompt, options):
        super().__init__(parent)
        self.title(title)
        self.geometry("480x360")
        self.resizable(False, False)
        self.grab_set()
        self.result = None
        ctk.CTkLabel(self, text=prompt, anchor="w").pack(padx=16, pady=(16,6), fill="x")
        self._listbox_var = ctk.StringVar()
        frame = ctk.CTkScrollableFrame(self, height=220)
        frame.pack(fill="both", expand=True, padx=16, pady=4)
        self._btns = []
        for i, opt in enumerate(options):
            b = ctk.CTkButton(frame, text=opt, anchor="w",
                              fg_color="transparent", hover_color="#3a3a3a",
                              command=lambda i=i: self._pick(i))
            b.pack(fill="x", pady=1)
            self._btns.append(b)
        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(fill="x", padx=16, pady=10)
        ctk.CTkButton(btn_row, text="Cancel", width=100,
                      fg_color=GRAY, hover_color=GRAY_HOVER,
                      command=self.destroy).pack(side="right")

    def _pick(self, idx):
        self.result = idx
        self.destroy()


# ================================================================
# ENTRY POINT
# ================================================================
def main():
    app = MeetingImportGUI()

    def on_close():
        try: app._save_config()
        except: pass
        try: app.quit()
        except: pass
        try: app.destroy()
        except: pass

    app.protocol("WM_DELETE_WINDOW", on_close)
    try:
        app.mainloop()
    except KeyboardInterrupt:
        on_close()

if __name__ == "__main__":
    main()
