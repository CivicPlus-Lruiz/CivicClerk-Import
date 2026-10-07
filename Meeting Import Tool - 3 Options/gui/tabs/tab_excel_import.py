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


# ============================================================================
# SUBTYPES AUTO-SAVE  — persist Subtypes tab edits to <output>_Config.json
# ============================================================================
# The read-only source only calls _auto_save_customer_config() from Sheet
# Mapping edits and Errors-tab fixes; nothing on the Subtypes tab saves. These
# wrappers add a debounced save after every subtype change without copying the
# original methods (so no re-sync burden if the source changes).
#
#   - Structural changes (scan / add / remove / clear / move) are wrapped:
#     run the original, then schedule a save.
#   - Field edits: after each _refresh_subtypes_display, walk the rebuilt rows
#     and add a trace to every entry/dropdown variable that schedules a save.
#     Traces are added after the vars are created, so a refresh (including the
#     one triggered by loading a config) never itself causes a save.
#   - Debounced via app.after so the original trace/command handlers have
#     already updated subtype_overrides before the save serializes it.

_SUBTYPE_SAVE_DELAY_MS = 400


def _install_subtype_autosave(proxy, app):
    state = {"after_id": None}

    def _do_save():
        state["after_id"] = None
        try:
            proxy._auto_save_customer_config()
        except Exception as e:
            print(f"[EI] Warning: subtype auto-save failed: {e}")

    def _schedule_save(*_a):
        if state["after_id"] is not None:
            try:
                app.after_cancel(state["after_id"])
            except Exception:
                pass
        state["after_id"] = app.after(_SUBTYPE_SAVE_DELAY_MS, _do_save)

    def _attach_traces(widget):
        for child in widget.winfo_children():
            var = getattr(child, "_textvariable", None) or getattr(child, "_variable", None)
            # Skip checkbox BooleanVars (row selection only; not persisted)
            if var is not None and child.__class__.__name__ in ("CTkEntry", "CTkOptionMenu"):
                try:
                    var.trace_add("write", _schedule_save)
                except Exception:
                    pass
            _attach_traces(child)

    orig_refresh = proxy._refresh_subtypes_display

    def _refresh_with_traces(self_p):
        orig_refresh()
        try:
            _attach_traces(self_p.subtypes_scroll)
        except Exception as e:
            print(f"[EI] Warning: could not attach subtype save traces: {e}")

    object.__setattr__(proxy, "_refresh_subtypes_display",
                       types.MethodType(_refresh_with_traces, proxy))

    def _wrap_structural(name):
        orig = getattr(proxy, name)

        def _wrapped(self_p, *args, **kwargs):
            result = orig(*args, **kwargs)
            _schedule_save()
            return result

        object.__setattr__(proxy, name, types.MethodType(_wrapped, proxy))

    for _name in ("_scan_subtypes", "_add_subtype_row", "_remove_selected_subtypes",
                  "_clear_subtypes", "_move_subtype_row"):
        _wrap_structural(_name)


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

    # ── Subtypes auto-save (must be installed before _build_subtypes_tab so
    #    the tab's buttons bind to the wrapped methods) ─────────────────────
    _install_subtype_autosave(proxy, app)

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
