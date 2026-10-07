"""
Excel Import Tab Module  — PROXY APPROACH
Loads Meeting_Import_GUI.py dynamically and builds its tabs directly into
the unified app's existing tab frames — no second window, no rewriting.
All UI code stays 100% in Meeting_Import_GUI.py exactly as authored.
"""

import importlib.util
import types
import sys
from pathlib import Path


# ============================================================================
# FTP FILENAME FIX  — view PDF, edit filename, rename on FTP + sync output sheet
# ============================================================================
# Feature: date-parse-failure rows (extra/missing digit, subtype/body typo) can
# be fixed WITHOUT leaving the app. User opens the PDF, edits the filename in a
# text box, and the tool renames the file on the FTP AND updates the identical
# name in the output sheet — because the importer requires byte-identical names
# on both sides.
#
# Design decisions (settled with the user):
#   - FTP-only. Azure (copy-then-delete) and flat structures are excluded.
#   - The user edits ONLY the filename; the full FTP path is discovered live by
#     re-listing the body folder (never shown, since users aren't technical).
#   - Ordering: validate -> collision-check -> confirm -> ftp.rename() FIRST,
#     then update the output sheet. If the sheet update fails after a successful
#     rename, surface a loud error (FTP is ahead; a rescan resyncs).
#   - Event Date column is updated ONLY if it was blank or wrong before AND the
#     new filename parses to exactly one date; otherwise the next dry run
#     re-flags it, so we never guess.

_OUTPUT_FILENAME_COLS = {
    # header text in the "CivicClerk Import" sheet -> nothing special, we match
    # by scanning these known filename-bearing headers for the old name.
    "Video File Name", "Agenda File Name", "Minutes File Name",
    "Agenda Packet File Name", "Close Caption File Name",
    "Notice File Name", "Other File Name",
}


def _ftp_connect(host, user, password):
    import ftplib
    ftp = ftplib.FTP()
    ftp.connect(host, 21, timeout=30)
    ftp.login(user or "anonymous", password or "")
    ftp.set_pasv(True)
    return ftp


def _ftp_find_file_dir(ftp, ftp_root, filename):
    """
    Locate the folder on the FTP that directly contains `filename`, by walking
    <ftp_root>/<body>/<type_folder>/. Returns the full remote dir path
    ('<ftp_root>/<body>/<type_folder>') or None if not found.

    Uses NLST per candidate folder (widely supported). One-level body scan,
    one-level type scan — matches the structure _do_ftp_scan expects.
    """
    def _names(path):
        try:
            raw = ftp.nlst(path)
            return [n.split("/")[-1] for n in raw if n.split("/")[-1] not in (".", "..")]
        except Exception:
            return []

    root = ftp_root.rstrip("/") or "/"
    for body in _names(root):
        body_path = root + "/" + body
        # A body entry could be a file at root (skip by catching nlst emptiness)
        for type_folder in _names(body_path):
            type_path = body_path + "/" + type_folder
            entries = _names(type_path)
            if filename in entries:
                return type_path
    return None


def _extract_single_date(proxy, filename):
    """Return 'MM/DD/YYYY' if the filename parses to exactly one date, else None."""
    try:
        dates = proxy._extract_dates_from_filename(filename)  # returns {'YYYY-MM-DD', ...}
    except Exception:
        return None
    if len(dates) != 1:
        return None
    ymd = next(iter(dates))  # 'YYYY-MM-DD'
    try:
        y, m, d = ymd.split("-")
        return f"{m}/{d}/{y}"
    except Exception:
        return None


def _sync_output_sheet(output_path, old_name, new_name):
    """
    In the output workbook, find the single cell holding old_name across the
    known filename columns and overwrite it with new_name.

    Returns (found: bool, row_idx: int|None, event_date_col: int|None,
             date_cell_value, header_map). Does NOT save — caller decides.
    """
    import openpyxl
    wb = openpyxl.load_workbook(output_path)
    ws = wb["CivicClerk Import"] if "CivicClerk Import" in wb.sheetnames else wb.active

    headers = {ws.cell(1, c).value: c for c in range(1, ws.max_column + 1)}
    fname_cols = [c for h, c in headers.items() if h in _OUTPUT_FILENAME_COLS]
    date_col = headers.get("Event Date")

    for r in range(2, ws.max_row + 1):
        for c in fname_cols:
            if ws.cell(r, c).value == old_name:
                return wb, ws, r, c, date_col
    return wb, ws, None, None, date_col


def _do_filename_fix(proxy, app, ctk, old_name, new_name, source_column):
    """
    Perform the full rename-and-sync. Returns (ok: bool, message: str).
    All destructive steps are guarded; caller has already confirmed.
    """
    from tkinter import messagebox

    host     = proxy.ftp_host.get().strip()
    user     = proxy.ftp_user.get().strip()
    password = proxy.ftp_password.get().strip()
    ftp_root = proxy.ftp_path.get().strip() or "/"

    if not host:
        return False, "No FTP host configured. Fill in FTP credentials on Setup."

    # ── 1. Connect + locate the file's folder (path hidden from user) ───────
    try:
        ftp = _ftp_connect(host, user, password)
    except Exception as e:
        return False, f"Could not connect to FTP:\n{e}"

    try:
        remote_dir = _ftp_find_file_dir(ftp, ftp_root, old_name)
        if remote_dir is None:
            return False, (f"Could not find '{old_name}' on the FTP.\n"
                           "It may have already been renamed or moved. "
                           "Try re-scanning.")

        old_path = remote_dir + "/" + old_name
        new_path = remote_dir + "/" + new_name

        # ── 2. Collision check — never rename onto an existing file ─────────
        try:
            existing = [n.split("/")[-1] for n in ftp.nlst(remote_dir)]
        except Exception:
            existing = []
        if new_name in existing:
            return False, (f"A file named '{new_name}' already exists in that "
                           "folder. Choose a different name.")

        # ── 3. Rename on FTP FIRST (source of truth for the importer) ───────
        try:
            ftp.rename(old_path, new_path)
        except Exception as e:
            return False, f"FTP rename failed:\n{e}\n\nNothing was changed."
    finally:
        try:
            ftp.quit()
        except Exception:
            pass

    # ── 4. Update the output sheet to the identical new name ────────────────
    output_path = proxy.output_file.get().strip()
    if not output_path or not Path(output_path).exists():
        return True, (f"Renamed on FTP to '{new_name}', but the output file "
                      "was not found, so the sheet was not updated. "
                      "Re-scan / re-run to resync.")

    try:
        wb, ws, row_idx, cell_col, date_col = _sync_output_sheet(
            output_path, old_name, new_name)
        if row_idx is None:
            wb.close()
            return True, (f"Renamed on FTP to '{new_name}'. The name was not "
                          "found in the output sheet (it may not have been "
                          "written yet) — a re-run will pick up the new name.")

        ws.cell(row_idx, cell_col).value = new_name

        # ── 5. Event Date: update only if blank/wrong AND new name parses ───
        if date_col:
            new_date = _extract_single_date(proxy, new_name)
            if new_date:
                cur = ws.cell(row_idx, date_col).value
                cur_str = str(cur).strip() if cur is not None else ""
                if not cur_str or cur_str != new_date:
                    ws.cell(row_idx, date_col).value = new_date

        wb.save(output_path)
        wb.close()
    except PermissionError:
        return False, (f"Renamed on FTP to '{new_name}', but the output file "
                       "is open in Excel and could not be updated.\n\n"
                       "Close it, then re-scan to resync the sheet.")
    except Exception as e:
        return False, (f"Renamed on FTP to '{new_name}', but updating the "
                       f"output sheet failed:\n{e}\n\n"
                       "The FTP is ahead of the sheet — re-scan to resync.")

    return True, f"Renamed to '{new_name}' on the FTP and updated the output sheet."


def _open_fix_name_dialog(proxy, app, ctk, old_name, source_column):
    """
    Download the PDF to a temp dir, open it in the PDFViewer, and show a text
    box pre-filled with the current filename for the user to edit.
    """
    import tempfile, os
    from tkinter import messagebox

    # Only meaningful for FTP source
    mode = proxy._source_mode.get()
    if mode != "ftp":
        messagebox.showinfo("Fix Filename",
                            "In-app filename fixing is available for FTP sources only.")
        return

    host     = proxy.ftp_host.get().strip()
    user     = proxy.ftp_user.get().strip()
    password = proxy.ftp_password.get().strip()
    ftp_root = proxy.ftp_path.get().strip() or "/"

    # ── Download the PDF to temp (best-effort; viewer is optional aid) ──────
    local_pdf = None
    if old_name.lower().endswith(".pdf"):
        try:
            ftp = _ftp_connect(host, user, password)
            try:
                remote_dir = _ftp_find_file_dir(ftp, ftp_root, old_name)
                if remote_dir:
                    tmp = Path(tempfile.gettempdir()) / f"ei_fix_{old_name}"
                    with open(tmp, "wb") as fh:
                        ftp.retrbinary(f"RETR {remote_dir}/{old_name}", fh.write)
                    local_pdf = str(tmp)
            finally:
                try:
                    ftp.quit()
                except Exception:
                    pass
        except Exception as e:
            print(f"[EI] Could not download PDF for preview: {e}")

    # ── Open the PDF viewer popup (real Tk root as parent) ──────────────────
    if local_pdf:
        try:
            mod = _load_meeting_import_gui_module()  # ensures pdf_viewer importable path
        except Exception:
            pass
        try:
            import pdf_viewer
            viewer = pdf_viewer.PDFViewer(app)
            viewer.open_file(local_pdf)
        except Exception as e:
            print(f"[EI] PDF viewer unavailable: {e}")

    # ── Edit dialog: single text box pre-filled with the current filename ───
    dlg = ctk.CTkToplevel(app)
    dlg.title("Fix Filename")
    dlg.geometry("560x220")
    dlg.transient(app)
    dlg.grab_set()

    ctk.CTkLabel(dlg, text="Edit the filename to fix the date, subtype, or body.",
                 font=ctk.CTkFont(size=13, weight="bold")).pack(padx=20, pady=(18, 4), anchor="w")
    ctk.CTkLabel(dlg, text="The file will be renamed on the FTP and the output "
                           "sheet updated to match.",
                 font=ctk.CTkFont(size=11), text_color="#aaaaaa",
                 wraplength=520, justify="left").pack(padx=20, pady=(0, 10), anchor="w")

    name_var = ctk.StringVar(value=old_name)
    entry = ctk.CTkEntry(dlg, textvariable=name_var, width=520,
                         font=ctk.CTkFont(size=12))
    entry.pack(padx=20, pady=(0, 6))
    entry.focus_set()

    status = ctk.CTkLabel(dlg, text="", font=ctk.CTkFont(size=11),
                          text_color="#ffaaaa")
    status.pack(padx=20, anchor="w")

    def _on_save():
        from tkinter import messagebox
        new_name = name_var.get().strip()
        if not new_name:
            status.configure(text="Filename cannot be empty.")
            return
        if new_name == old_name:
            status.configure(text="Filename is unchanged.")
            return
        if any(ch in new_name for ch in '/\\:*?"<>|'):
            status.configure(text="Filename contains invalid characters.")
            return

        # Validate it parses to a single date (warn but allow — dry run re-checks)
        parsed = _extract_single_date(proxy, new_name)
        warn = ""
        if parsed is None:
            warn = ("\n\nNote: the new name still doesn't parse to a single "
                    "clear date. The rename will proceed, but the next dry run "
                    "may re-flag it.")

        if not messagebox.askyesno(
            "Confirm Rename",
            f"Rename on the FTP:\n\n"
            f"  {old_name}\n      ↓\n  {new_name}\n\n"
            f"This will also update the output sheet.{warn}\n\nProceed?"):
            return

        ok, msg = _do_filename_fix(proxy, app, ctk, old_name, new_name, source_column)
        if ok:
            messagebox.showinfo("Rename Complete", msg)
            dlg.destroy()
        else:
            messagebox.showerror("Rename Failed", msg)
            status.configure(text="Rename failed — see message above.")

    btn_row = ctk.CTkFrame(dlg, fg_color="transparent")
    btn_row.pack(side="bottom", fill="x", padx=20, pady=14)
    ctk.CTkButton(btn_row, text="Cancel", width=100, fg_color="#3a3a3a",
                  hover_color="#4a4a4a", command=dlg.destroy).pack(side="right", padx=(8, 0))
    ctk.CTkButton(btn_row, text="Rename + Sync", width=140, fg_color="#2e7d32",
                  hover_color="#1b5e20", font=ctk.CTkFont(size=13, weight="bold"),
                  command=_on_save).pack(side="right")


def _inject_fix_name_buttons(proxy, app, ctk):
    """
    Wrap _populate_errors_tab so that after the source builds the date-error
    rows, we add a 'Fix Name' button into each row. Survives re-populations.
    FTP-only: buttons are added but disabled when source mode isn't FTP.
    """
    _orig_populate = getattr(proxy, "_populate_errors_tab", None)
    if _orig_populate is None:
        return

    def _wrapped_populate(*a, **k):
        _orig_populate(*a, **k)
        try:
            _decorate_date_rows()
        except Exception as e:
            print(f"[EI] Could not add Fix Name buttons: {e}")

    def _decorate_date_rows():
        scroll = getattr(proxy, "date_errors_scroll", None)
        if scroll is None:
            return
        is_ftp = (proxy._source_mode.get() == "ftp")
        # Each child is a row frame with a 4-col grid (0-3). We add col 4.
        # Recover the filename from the row's first label (col 0).
        for row in scroll.winfo_children():
            # Skip if we already added a button to this row
            if getattr(row, "_ei_has_fix_btn", False):
                continue
            fname = None
            src_col = ""
            for child in row.winfo_children():
                info = child.grid_info()
                if info.get("column") == 0 and hasattr(child, "cget"):
                    try:
                        fname = child.cget("text")
                    except Exception:
                        pass
                if info.get("column") == 1 and hasattr(child, "cget"):
                    try:
                        src_col = child.cget("text")
                    except Exception:
                        pass
            if not fname:
                continue
            try:
                row.grid_columnconfigure(4, minsize=110)
                btn = ctk.CTkButton(
                    row, text="Fix Name", width=100, height=26,
                    font=ctk.CTkFont(size=11, weight="bold"),
                    fg_color="#2e7d32" if is_ftp else "#3a3a3a",
                    hover_color="#1b5e20" if is_ftp else "#3a3a3a",
                    state="normal" if is_ftp else "disabled",
                    command=lambda fn=fname, sc=src_col:
                        _open_fix_name_dialog(proxy, app, ctk, fn, sc),
                )
                btn.grid(row=0, column=4, padx=6, pady=5, sticky="e")
                object.__setattr__(row, "_ei_has_fix_btn", True)
            except Exception as e:
                print(f"[EI] row decorate failed: {e}")

    object.__setattr__(proxy, "_populate_errors_tab",
                       types.MethodType(lambda self_p, *a, **k: _wrapped_populate(*a, **k), proxy))


def _load_meeting_import_gui_module():
    """
    Dynamically load Meeting_Import_GUI.py from the same directory as main.py.
    Returns the module, or None if not found.
    """
    candidates = [
        Path(sys.argv[0]).parent / "Meeting_Import_GUI.py",
        Path.cwd() / "Meeting_Import_GUI.py",
    ]
    for path in candidates:
        if path.exists():
            spec   = importlib.util.spec_from_file_location("Meeting_Import_GUI", path)
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            return module
    return None


def _inject_errors_rescan_button(proxy, app, ctk):
    """
    Add a 'Re-scan + Dry Run' button to the top of the EI Errors tab.

    Enabled only when the data source is FTP or Azure; greyed out (with an
    explanatory tooltip-style label) in Excel mode. Stays in sync with the
    _source_mode variable so switching source on the Setup tab updates it live.
    """
    frame = app.tab_frames.get("EI Errors")
    if frame is None:
        return

    # A slim bar pinned to the very top of the Errors frame, above the
    # source-built header (which is packed after this in build order — but the
    # source content is already packed, so we pack this bar and lift it to top).
    bar = ctk.CTkFrame(frame, fg_color="transparent")
    bar.pack(side="top", fill="x", padx=20, pady=(8, 0))

    btn = ctk.CTkButton(
        bar,
        text="\u21ba  Re-scan + Dry Run",
        width=200,
        height=32,
        font=ctk.CTkFont(size=13, weight="bold"),
        fg_color="#2e7d32",
        hover_color="#1b5e20",
        command=proxy._errors_tab_rescan,
    )
    btn.pack(side="right")

    hint = ctk.CTkLabel(
        bar,
        text="",
        font=ctk.CTkFont(size=11),
        text_color="#888888",
    )
    hint.pack(side="right", padx=(0, 12))

    # Store on proxy so the rescan handlers can re-enable it
    object.__setattr__(proxy, "_ei_errors_rescan_btn", btn)

    def _sync(*_a):
        mode = proxy._source_mode.get()
        if mode in ("ftp", "azure"):
            btn.configure(state="normal",
                          fg_color="#2e7d32", hover_color="#1b5e20")
            hint.configure(text=f"Source: {mode.upper()}")
        else:
            btn.configure(state="disabled",
                          fg_color="#3a3a3a", hover_color="#3a3a3a")
            hint.configure(text="Re-scan available for FTP / Azure sources only")

    _sync()
    try:
        proxy._source_mode.trace_add("write", _sync)
    except Exception:
        # Older Tk: fall back to trace()
        try:
            proxy._source_mode.trace("w", _sync)
        except Exception:
            pass


def build_excel_import_tabs(app):
    """
    Called once from meeting_import_app._setup_ui().

    Loads Meeting_Import_GUI, initialises all its state vars on `app` (prefixed
    ei_), then builds every tab directly into the unified app's EI tab frames
    via a lightweight proxy object — no second Tk window created.
    """
    import customtkinter as ctk

    mod = _load_meeting_import_gui_module()

    if mod is None:
        # Graceful fallback — show a message in each EI frame
        for tab_name in ["EI Setup", "EI Sheet Mapping", "EI Subtypes",
                         "EI Run", "EI Errors", "EI Results"]:
            frame = app.tab_frames.get(tab_name)
            if frame:
                ctk.CTkLabel(
                    frame,
                    text="⚠  Meeting_Import_GUI.py not found.\n\n"
                         "Place it in the same folder as main.py to enable Excel Import.",
                    font=ctk.CTkFont(size=14),
                    text_color="#ff9800",
                    justify="center",
                ).place(relx=0.5, rely=0.5, anchor="center")
        app.ei_proxy = None
        return

    # ── Initialise all state vars from MeetingImportGUI.__init__ onto app ───
    # Prefixed ei_ to avoid collisions with existing app state.
    PATTERN_PRESETS = mod.PATTERN_PRESETS

    app.ei_input_file        = ctk.StringVar()
    app.ei_output_file       = ctk.StringVar()
    app.ei_tab_configs       = []
    app.ei_subtype_overrides = []
    app.ei_last_result       = None
    app.ei_running           = False

    app.ei_ftp_host          = ctk.StringVar()
    app.ei_ftp_user          = ctk.StringVar()
    app.ei_ftp_password      = ctk.StringVar()
    app.ei_ftp_path          = ctk.StringVar(value="/")
    app.ei_az_sas_url        = ctk.StringVar()
    app.ei_az_parsed_path    = ctk.StringVar()
    app.ei_scan_pattern      = ctk.StringVar(value=PATTERN_PRESETS[0])
    app.ei_scan_sep          = ctk.StringVar(value="_")
    app.ei_scan_custom       = ctk.StringVar()
    app.ei_scan_filter       = ctk.StringVar(value="All Meeting Bodies")
    app.ei_ftp_structure     = ctk.StringVar(value="None")
    app.ei_az_structure      = ctk.StringVar(value="None")
    app.ei_flat_pattern      = ctk.StringVar(value=PATTERN_PRESETS[0])
    app.ei_flat_sep          = ctk.StringVar(value="_")
    app.ei_flat_custom       = ctk.StringVar()
    app.ei_date_fixes        = {}
    app.ei_dup_fixes         = {}
    app.ei_body_fixes        = {}
    app.ei_pending_errors    = []
    app.ei_source_mode       = ctk.StringVar(value="excel")

    # ── Attribute map: original name → ei_-prefixed name on app ────────────
    _ATTR_MAP = {
        "input_file":         "ei_input_file",
        "output_file":        "ei_output_file",
        "tab_configs":        "ei_tab_configs",
        "subtype_overrides":  "ei_subtype_overrides",
        "last_result":        "ei_last_result",
        "_running":           "ei_running",
        "ftp_host":           "ei_ftp_host",
        "ftp_user":           "ei_ftp_user",
        "ftp_password":       "ei_ftp_password",
        "ftp_path":           "ei_ftp_path",
        "az_sas_url":         "ei_az_sas_url",
        "az_parsed_path":     "ei_az_parsed_path",
        "scan_pattern":       "ei_scan_pattern",
        "scan_sep":           "ei_scan_sep",
        "scan_custom":        "ei_scan_custom",
        "scan_filter":        "ei_scan_filter",
        "ftp_structure":      "ei_ftp_structure",
        "az_structure":       "ei_az_structure",
        "flat_pattern":       "ei_flat_pattern",
        "flat_sep":           "ei_flat_sep",
        "flat_custom":        "ei_flat_custom",
        "date_fixes":         "ei_date_fixes",
        "dup_fixes":          "ei_dup_fixes",
        "body_fixes":         "ei_body_fixes",
        "_pending_errors":    "ei_pending_errors",
        "_source_mode":       "ei_source_mode",
    }

    # ── Proxy class ─────────────────────────────────────────────────────────
    class Proxy:
        """
        Looks like a MeetingImportGUI instance to all the _build_*_tab methods,
        but delegates state to `app` and tab frames to the unified app's EI frames.
        """

        # tab_frames: map original names → unified app's EI-prefixed frames
        @property
        def tab_frames(self_p):
            return {
                "Setup":         app.tab_frames["EI Setup"],
                "Sheet Mapping": app.tab_frames["EI Sheet Mapping"],
                "Subtypes":      app.tab_frames["EI Subtypes"],
                "Run":           app.tab_frames["EI Run"],
                "Errors":        app.tab_frames["EI Errors"],
                "Results":       app.tab_frames["EI Results"],
            }

        # tab_btns: the tab builders assign to this; we just store locally
        # (the unified app's tab bar handles actual button highlighting)
        @property
        def tab_btns(self_p):
            if not hasattr(self_p, "_ei_tab_btns"):
                object.__setattr__(self_p, "_ei_tab_btns", {})
            return self_p._ei_tab_btns

        @property
        def current_tab(self_p):
            # Strip "EI " prefix so _switch_tab comparisons work
            ct = app.current_tab
            return ct[3:] if ct.startswith("EI ") else ct

        @current_tab.setter
        def current_tab(self_p, value):
            pass  # managed by unified app

        def _switch_tab(self_p, name):
            # Explicitly hide every EI frame before showing the target one.
            # This prevents ghosting caused by multiple EI frames being packed
            # inside the unified app's content_frame simultaneously.
            for ei_name in ["EI Setup", "EI Sheet Mapping", "EI Subtypes",
                            "EI Run", "EI Errors", "EI Results"]:
                try:
                    app.tab_frames[ei_name].pack_forget()
                except Exception:
                    pass
            app._switch_tab("EI " + name)

        # Tkinter/CTk root methods delegated to the real root
        def after(self_p, ms, func=None, *args):
            return app.after(ms, func, *args)

        def update(self_p):
            app.update()

        def __getattr__(self_p, name):
            if name in _ATTR_MAP:
                return getattr(app, _ATTR_MAP[name])
            # Forward class-level constants (e.g. _AGENDA_NAMES) and any other
            # attributes that live on the real MeetingImportGUI class or on app
            # (e.g. .tk, .after) so threads don't crash when accessing them.
            try:
                return getattr(gui_cls, name)
            except AttributeError:
                pass
            try:
                return getattr(app, name)
            except AttributeError:
                pass
            raise AttributeError(f"Proxy has no attribute '{name}'")

        def __setattr__(self_p, name, value):
            if name in _ATTR_MAP:
                setattr(app, _ATTR_MAP[name], value)
            else:
                object.__setattr__(self_p, name, value)

    proxy = Proxy()

    # ── Bind all MeetingImportGUI methods onto the proxy ────────────────────
    gui_cls = mod.MeetingImportGUI

    # ── "Notice" folder support ──────────────────────────────────────────────
    # Meeting_Import_GUI.py has no _NOTICE_NAMES set, so a folder named
    # Notice/Notices fell into `unrecognized` in _do_ftp_scan and its files
    # were dropped before anything else ran. Added here as a class attribute
    # so it sits alongside the read-only file's own _AGENDA_NAMES/_MINUTES_NAMES/
    # etc. and is referenced the same way (self._NOTICE_NAMES) from the
    # single-pass override below.
    gui_cls._NOTICE_NAMES = {"notice", "public notices", "notices"}

    _skip = {
        "__init__", "__class__", "__new__", "_build_ui", "mainloop",
        "destroy", "quit", "title", "geometry", "__dict__",
        "__weakref__", "__doc__",
        "_switch_tab",   # Proxy defines its own — must not be overwritten by binding loop
    }
    for name in dir(gui_cls):
        if name in _skip:
            continue
        try:
            fn = getattr(gui_cls, name)
            if callable(fn) and isinstance(fn, types.FunctionType):
                object.__setattr__(proxy, name, types.MethodType(fn, proxy))
        except Exception:
            pass

    # ── Build preview/error trees WITHOUT calling theme_use("clam") ─────────
    # The original methods call style.theme_use("clam") which is app-wide and
    # would override the dark theme on every other tab.  These replacements do
    # everything the originals do except that one call.
    def _safe_build_preview_tree(self_p):
        import tkinter.ttk as ttk
        style = ttk.Style()
        # deliberately skip: style.theme_use("clam")
        style.configure("Custom.Treeview",
                        background="#2b2b2b", foreground="white",
                        fieldbackground="#2b2b2b", rowheight=24,
                        font=("Arial", 11))
        style.configure("Custom.Treeview.Heading",
                        background="#1f538d", foreground="white",
                        font=("Arial", 11, "bold"))
        style.map("Custom.Treeview", background=[("selected", "#1f538d")])

        cols = ["Event Name", "Event Date", "Event Category", "Meeting Type",
                "Agenda", "Minutes", "Video", "Packet"]
        self_p.preview_tree = ttk.Treeview(self_p.preview_outer, columns=cols,
                                            show="headings", style="Custom.Treeview")
        widths = [260, 90, 190, 120, 55, 55, 55, 55]
        for col, w in zip(cols, widths):
            self_p.preview_tree.heading(col, text=col)
            self_p.preview_tree.column(col, width=w, minwidth=40)
        vsb = ttk.Scrollbar(self_p.preview_outer, orient="vertical",
                            command=self_p.preview_tree.yview)
        self_p.preview_tree.configure(yscrollcommand=vsb.set)
        self_p.preview_tree.pack(side="left", fill="both", expand=True, padx=4, pady=4)
        vsb.pack(side="right", fill="y", pady=4)
        self_p.preview_tree.tag_configure("dup", background="#5a4a00")

    def _safe_build_error_tree(self_p):
        import tkinter.ttk as ttk
        cols = ["Filename", "Source", "Issue"]
        self_p.error_tree = ttk.Treeview(self_p.error_outer, columns=cols,
                                          show="headings", style="Custom.Treeview")
        widths = [380, 90, 500]
        for col, w in zip(cols, widths):
            self_p.error_tree.heading(col, text=col)
            self_p.error_tree.column(col, width=w, minwidth=50)
        vsb2 = ttk.Scrollbar(self_p.error_outer, orient="vertical",
                              command=self_p.error_tree.yview)
        self_p.error_tree.configure(yscrollcommand=vsb2.set)
        self_p.error_tree.pack(side="left", fill="both", expand=True, padx=4, pady=4)
        vsb2.pack(side="right", fill="y", pady=4)

    object.__setattr__(proxy, "_build_preview_tree",
                       types.MethodType(_safe_build_preview_tree, proxy))
    object.__setattr__(proxy, "_build_error_tree",
                       types.MethodType(_safe_build_error_tree, proxy))

    # ── Override _scan_output_folder to use real app window as CTkToplevel parent
    # The original passes `self` (Proxy) to _PickDialog, which inherits CTkToplevel.
    # CTkToplevel's __init__ calls root.title() internally — this fails because
    # the Proxy is not a real Tk widget.  We replace self→app for that one call.
    def _safe_scan_output_folder(self_p, mod=mod, app=app):
        from pathlib import Path
        from tkinter import filedialog, messagebox
        op = self_p.output_file.get().strip()
        folder = str(Path(op).parent) if op else ""
        folder = filedialog.askdirectory(title="Select folder to scan", initialdir=folder or ".")
        if not folder:
            return
        patterns = ["*_FTP_Input.xlsx", "*_Azure_Input.xlsx", "*_Output.xlsx",
                    "*_Import.xlsx", "*.xlsx"]
        found_files = []
        for pat in patterns:
            found_files.extend(Path(folder).glob(pat))
        found_files = sorted(set(found_files), key=lambda p: p.name)
        if not found_files:
            messagebox.showinfo("Scan Folder", f"No Excel files found in:\n{folder}")
            return
        names = [f.name for f in found_files]
        # Pass `app` (real Tk window) as parent so CTkToplevel initialises correctly
        sel = mod._PickDialog(app, "Select Input File", "Choose a file to use as input:", names)
        app.wait_window(sel)
        if sel.result is not None:
            chosen = found_files[sel.result]
            self_p.input_file.set(str(chosen))
            self_p._scan_sheets_silent()
            messagebox.showinfo("File Set", f"Input file set to:\n{chosen.name}")

    object.__setattr__(proxy, "_scan_output_folder",
                       types.MethodType(_safe_scan_output_folder, proxy))

    # ── Errors-tab "Re-scan + Dry Run" chaining ─────────────────────────────
    # _rescan() runs in a background thread and returns immediately, so we can't
    # call _rescan() then a dry run sequentially. Instead we wrap _on_rescan_done
    # (which fires on the main thread once the input Excel is rebuilt) so that,
    # WHEN triggered from the Errors tab, it chains into a dry run automatically.
    # A one-shot flag scopes this so the Run-tab Re-scan button keeps its old
    # behavior (rebuild only, no auto dry run).
    _orig_on_rescan_done = gui_cls._on_rescan_done

    def _wrapped_on_rescan_done(self_p, result, mode):
        _orig_on_rescan_done(self_p, result, mode)
        if getattr(self_p, "_ei_chain_dryrun", False):
            object.__setattr__(self_p, "_ei_chain_dryrun", False)
            # input Excel is now rebuilt and self.input_file is set — run dry run
            self_p._errors_tab_dry_run()

    object.__setattr__(proxy, "_on_rescan_done",
                       types.MethodType(_wrapped_on_rescan_done, proxy))

    # Also re-enable the Errors-tab button on rescan error (the original
    # _on_rescan_error only re-enables the Run-tab rescan_btn).
    _orig_on_rescan_error = gui_cls._on_rescan_error

    def _wrapped_on_rescan_error(self_p, msg, tb):
        object.__setattr__(self_p, "_ei_chain_dryrun", False)
        _orig_on_rescan_error(self_p, msg, tb)
        btn = getattr(self_p, "_ei_errors_rescan_btn", None)
        if btn is not None:
            try:
                btn.configure(state="normal", text="\u21ba  Re-scan + Dry Run")
            except Exception:
                pass

    object.__setattr__(proxy, "_on_rescan_error",
                       types.MethodType(_wrapped_on_rescan_error, proxy))

    def _errors_tab_rescan(self_p):
        """Errors-tab action: rebuild input from FTP/Azure, then auto dry-run."""
        mode = self_p._source_mode.get()
        if mode not in ("ftp", "azure"):
            return  # button is disabled in excel mode; guard anyway
        object.__setattr__(self_p, "_ei_chain_dryrun", True)
        btn = getattr(self_p, "_ei_errors_rescan_btn", None)
        if btn is not None:
            try:
                btn.configure(state="disabled", text="Scanning…")
            except Exception:
                pass
        self_p._rescan()

    object.__setattr__(proxy, "_errors_tab_rescan",
                       types.MethodType(_errors_tab_rescan, proxy))

    # ── Single-pass "Notice" folder recognition ──────────────────────────────
    # Full overrides of _do_ftp_scan, _on_ftp_done, and _execute_import.
    # "Notice File Name" is a real, separate column in the output schema, but
    # _process_col/col_map live as closures inside _execute_import, so there's
    # no way to add just the notice branch without copying the whole method —
    # same story for the folder-classification chain in _do_ftp_scan and the
    # intermediate-Excel column building in _on_ftp_done. These are full
    # copies of the originals with a `notices` bucket added at each stage,
    # classified in the same single pass as Agenda/Minutes/Video/etc. via
    # self._NOTICE_NAMES (set on gui_cls above).
    def _do_ftp_scan_with_notices(self, host: str, user: str, password: str, ftp_root: str) -> dict:
        import ftplib

        debug = []

        ftp = ftplib.FTP()
        ftp.connect(host, 21, timeout=30)
        welcome = ftp.login(user, password)
        debug.append(f"Login OK: {welcome.strip()}")
        ftp.set_pasv(True)

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
            rec = {"body_name": body_name, "agendas": [], "minutes": [], "videos": [],
                   "packets": [], "other": [], "cc": [], "notices": []}

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
                elif key_lower in self._NOTICE_NAMES:
                    col = "notices"
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

    object.__setattr__(proxy, "_do_ftp_scan",
                       types.MethodType(_do_ftp_scan_with_notices, proxy))

    def _on_ftp_done_with_notices(self, result: dict, mod=mod):
        from tkinter import messagebox
        import openpyxl

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

        output_path = self.output_file.get().strip()
        p = Path(output_path)
        input_path = str(p.parent / (p.stem + "_FTP_Input.xlsx"))

        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        for body in bodies:
            full_name  = body["body_name"]
            sheet_name = full_name[:31]
            ws = wb.create_sheet(title=sheet_name)
            ws.append(["Agendas", "Minutes", "Videos", "Packets", "Other", "Close Captions", "Notices"])
            cc_files = body.get("cc", [])
            max_rows = max(len(body["agendas"]), len(body["minutes"]),
                          len(body["videos"]), len(body.get("packets",[])),
                          len(body.get("other",[])), len(body.get("notices",[])), 1)
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
                    body.get("notices",[])[i]       if i < len(body.get("notices",[]))      else "",
                ])

        wb.save(input_path)
        self.input_file.set(input_path)

        existing_sheets = {tc.sheet_name for tc in self.tab_configs}
        added = 0
        for body in bodies:
            full_name  = body["body_name"]
            sheet_name = full_name[:31]
            if sheet_name not in existing_sheets:
                self.tab_configs.append(mod.TabConfig(
                    sheet_name=sheet_name,
                    body_name=full_name,
                ))
                added += 1

        self._refresh_mapping_display()

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

    object.__setattr__(proxy, "_on_ftp_done",
                       types.MethodType(_on_ftp_done_with_notices, proxy))

    def _execute_import_with_notices(self, dry_run: bool = False, mod=mod) -> dict:
        import pandas as pd
        from openpyxl import load_workbook
        from openpyxl.styles import PatternFill, Font, Alignment
        from openpyxl.utils import get_column_letter

        input_path  = self.input_file.get().strip()
        output_path = self.output_file.get().strip()

        engine = self._load_engine()
        if not engine:
            raise RuntimeError("Engine not found.")

        tab_body_map = {
            tc.sheet_name.strip(): tc
            for tc in self.tab_configs if tc.sheet_name.strip()
        }

        body_name_map = {
            tc.body_name.strip().lower(): tc
            for tc in self.tab_configs if tc.body_name.strip()
        }

        _body_specific  = []
        _universal      = []
        _body_catchall  = {}
        _univ_catchall  = None

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
                _univ_catchall = so

        _body_specific.sort(key=lambda x: max(len(a) for a in x[1]), reverse=True)
        _universal.sort(key=lambda x: max(len(a) for a in x[0]), reverse=True)

        def _match_alias(body: str, fname: str):
            fname_lower = fname.lower()
            body_lower  = body.lower()
            for row_body, aliases, so in _body_specific:
                if row_body != body_lower:
                    continue
                for alias in aliases:
                    if alias in fname_lower:
                        return so
            for aliases, so in _universal:
                for alias in aliases:
                    if alias in fname_lower:
                        return so
            if body_lower in _body_catchall:
                return _body_catchall[body_lower]
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

        all_meetings    = {}
        all_errors      = []
        keys_with_dupes = set()
        seen            = {}

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
                elif cl in ("notice","notices","public notices"):                   col_map["notice"]  = col

            def _process_col(col_key, file_field):
                nonlocal forced_body
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

                    fname          = mod._extract_filename(raw)
                    date_fmt_pref  = (forced_body.date_format.strip()
                                      if forced_body and forced_body.date_format.strip()
                                      else "MM-DD-YYYY")

                    if fname in self.date_fixes:
                        fmt_str  = self.date_fixes[fname]
                        fixed_dt = mod._parse_with_format(fname, fmt_str)
                        if fixed_dt:
                            parsed = {"date": fixed_dt, "is_error": False, "error_reason": "",
                                      "body": None, "subtype": "", "match_key": None,
                                      "stored_value": fname}
                        else:
                            parsed = engine.parse_filename(fname, date_fmt_pref)
                    else:
                        parsed = engine.parse_filename(fname, date_fmt_pref)

                    if fname in self.body_fixes:
                        fixed_sheet = self.body_fixes[fname]
                        forced_body = tab_body_map.get(fixed_sheet) or forced_body

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
                        _cfg      = body_name_map.get(body.strip().lower())
                        base_mt   = _cfg.meeting_type.strip()   if _cfg else ""
                        base_ec   = _cfg.event_category.strip() if _cfg else ""
                        base_time = _cfg.default_time           if _cfg else ""

                    ovr = _match_alias(body, fname)
                    if ovr:
                        subtype    = ovr.subtype.strip()
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
                        fix_key  = (existing, fname, file_field)
                        fix_key2 = (fname, existing, file_field)
                        fix = self.dup_fixes.get(fix_key) or self.dup_fixes.get(fix_key2)

                        if fix in ("other1",):
                            rec["Other File Name"] = existing
                            rec[file_field]        = fname
                        elif fix in ("other2",):
                            rec["Other File Name"] = fname
                        elif fix == "keep2":
                            rec[file_field] = fname
                        else:
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
            _process_col("notice",  "Notice File Name")

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

        output_p    = Path(output_path)
        errors_path = str(output_p.parent / (output_p.stem + "_Errors.xlsx"))

        self._log(f"\n✓  {len(df_out)} meeting rows")
        self._log(f"✓  {len(df_errors)} errors / flags")

        if dry_run:
            self._log("\n🔍  DRY RUN complete — no file written.")
        else:
            self._log("Writing output…")

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

    object.__setattr__(proxy, "_execute_import",
                       types.MethodType(_execute_import_with_notices, proxy))

    # ── Build all tabs ───────────────────────────────────────────────────────
    proxy._build_setup_tab()
    proxy._build_mapping_tab()
    proxy._build_subtypes_tab()
    proxy._build_run_tab()
    proxy._build_errors_tab()
    proxy._build_results_tab()

    # ── Inject "Re-scan + Dry Run" button into the Errors tab ────────────────
    # The Errors tab is fully built by the read-only source; we add our button
    # into its existing frame afterward. It is greyed out unless the data source
    # is FTP or Azure, and reflects source-mode changes made on the Setup tab.
    _inject_errors_rescan_button(proxy, app, ctk)

    # ── Inject "Fix Name" buttons into date-error rows (FTP rename + sync) ───
    _inject_fix_name_buttons(proxy, app, ctk)

    # ── Restore saved config (tab_configs + subtype_overrides + paths) ───────
    # _load_config is normally called from MeetingImportGUI.__init__, which we
    # skip in the proxy approach.  Call it explicitly now that all tabs are built
    # so _refresh_mapping_display / _refresh_subtypes_display can run safely.
    try:
        proxy._load_config()
    except Exception as e:
        print(f"[EI] Warning: could not load saved config: {e}")

    # Store proxy on app for config persistence
    app.ei_proxy = proxy
