"""
Meeting Import GUI v1.5 - Main Application
Thin orchestration layer: window setup, shared state, tab routing.
All tab logic lives in gui/tabs/.
"""

import customtkinter as ctk
from tkinter import messagebox
import json
import threading
from pathlib import Path
from typing import Optional, List, Dict, Any

from gui.models import MeetingBody, Subtype, FileError
from gui.constants import CHECKLIST_ITEMS

# SharePoint / backend (optional)
try:
    from sharepoint_integration import SharePointManager
    from backend_wrapper import BackendWrapper
except ImportError as e:
    print(f"Warning: Could not import optional modules: {e}")
    SharePointManager = None
    BackendWrapper = None

# Tab modules
from gui.tabs.tab_import_type import build_import_type_tab, select_import_type
from gui.tabs.tab_excel_import import build_excel_import_tabs
from gui.tabs.tab_checklist import (
    build_checklist_tab, build_checklist_items,
    open_link, on_siteid_changed, save_checklist_state,
)
from gui.tabs.tab_meeting_bodies import (
    build_meeting_bodies_tab, browse_base_dir, import_from_sheet,
    auto_scan_meetings, add_meeting_body, remove_selected_bodies,
    clear_all_meeting_bodies, refresh_meeting_bodies_display,
    update_meeting_time, browse_body_folder,
)
from gui.tabs.tab_subtypes import (
    build_subtypes_tab, create_advanced_matching_section,
    update_scan_filter_options, toggle_advanced_matching,
    on_pattern_selected, auto_scan_subtypes,
    next_from_meeting_bodies, next_from_subtypes,
    add_subtype, remove_selected_subtypes, clear_all_subtypes,
    refresh_subtypes_display, update_subtype_time,
    move_subtype, drag_release,
)
from gui.tabs.tab_file_errors import (
    build_file_errors_tab, refresh_file_errors_display,
    run_ocr_on_selected, apply_ocr_suggestion, edit_ocr_suggestion,
    rename_selected_errors, rename_single_file, bulk_rename_files,
    apply_bulk_rename, open_error_location, mark_errors_resolved,
    remove_duplicates, select_all_errors, deselect_all_errors, rerun_import,
    preview_selected_pdf,
)
from gui.tabs.tab_run import (
    build_run_tab, run_dry_run, run_import, execute_import,
    mock_import, display_results, populate_file_errors,
    open_review_folders, export_excel,
)
from gui.tabs.tab_ac_event_name import (
    build_ac_event_name_tab, ac_browse_file,
    ac_populate_event_name_rows, ac_next_to_meeting_bodies,
)
from gui.tabs.tab_ac_meeting_bodies import (
    build_ac_meeting_bodies_tab, ac_preload_meeting_bodies,
    ac_save_time, ac_run_import,
)


class MeetingImportGUI(ctk.CTk):
    """Main application window — wires together all tab modules."""

    def __init__(self):
        super().__init__()

        # ── Window setup ──────────────────────────────────────────────
        self.title("Meeting Import Tool - CivicClerk")
        self.geometry("1600x900")
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # ── Shared state ──────────────────────────────────────────────
        self.site_id = ctk.StringVar(value="")
        self.base_dir = ctk.StringVar(value="")
        self.date_format = ctk.StringVar(value="US DATE")
        self.import_type: Optional[str] = None  # "standard" | "agenda_center"

        # Standard Import state
        self.meeting_bodies: List[MeetingBody] = []
        self.subtypes: List[Subtype] = []
        self.file_errors: List[FileError] = []
        self.checklist_state: Dict[str, Dict[int, bool]] = {}
        self.last_result: Optional[Dict[str, Any]] = None

        # Agenda Center state
        self.ac_source_file: Optional[str] = None
        self.ac_df = None
        self.ac_event_name_edits: Dict[int, str] = {}
        self.ac_category_renames: Dict[str, str] = {}
        self.ac_type_renames: Dict[str, str] = {}
        self.ac_default_times: Dict[str, str] = {}
        self.ac_time_hour_vars: Dict[str, Any] = {}
        self.ac_time_min_vars: Dict[str, Any] = {}

        # Managers (initialized after UI)
        self.sharepoint_manager = None
        self.backend_wrapper = None

        # ── Build UI & initialize ─────────────────────────────────────
        self._setup_ui()
        self.after(100, self._initialize_managers)

    # ================================================================
    # UI SETUP & TAB ROUTING
    # ================================================================

    def _setup_ui(self):
        """Build the shell: tab bar + content area, then delegate to tab modules."""
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        tab_bar = ctk.CTkFrame(main_frame, fg_color="#2b2b2b", height=60)
        tab_bar.pack(fill="x", pady=(0, 10))
        tab_bar.pack_propagate(False)

        self.button_container = ctk.CTkFrame(tab_bar, fg_color="transparent")
        self.button_container.place(relx=0.5, rely=0.5, anchor="center")

        self.ALL_TABS = [
            "Import Type", "Checklist", "Meeting Bodies", "Subtypes",
            "Run", "File Errors", "AC Event Name", "AC Meeting Bodies",
            "EI Setup", "EI Sheet Mapping", "EI Subtypes",
            "EI Run", "EI Errors", "EI Results",
        ]
        self.TABS_BY_TYPE = {
            "standard":      ["Import Type", "Checklist", "Meeting Bodies", "Subtypes", "Run", "File Errors"],
            "agenda_center": ["Import Type", "AC Event Name", "AC Meeting Bodies"],
            "excel_import":  ["Import Type", "EI Setup", "EI Sheet Mapping", "EI Subtypes", "EI Run", "EI Errors", "EI Results"],
        }
        self.tab_buttons: Dict[str, ctk.CTkButton] = {}
        self.tab_frames: Dict[str, ctk.CTkFrame] = {}

        self.content_frame = ctk.CTkFrame(main_frame)
        self.content_frame.pack(fill="both", expand=True)

        for tab_name in self.ALL_TABS:
            self.tab_frames[tab_name] = ctk.CTkFrame(self.content_frame)

        # Delegate each tab build to its module
        self._build_import_type_tab()
        self._build_checklist_tab()
        self._build_meeting_bodies_tab()
        self._build_subtypes_tab()
        self._build_file_errors_tab()
        self._build_run_tab()
        self._build_ac_event_name_tab()
        self._build_ac_meeting_bodies_tab()
        build_excel_import_tabs(self)

        self.tabs = ["Import Type"]
        self.current_tab = "Import Type"
        self._rebuild_tab_buttons()
        self.tab_frames["Import Type"].pack(fill="both", expand=True)

    def _rebuild_tab_buttons(self):
        """Recreate tab buttons to match current self.tabs list."""
        for widget in self.button_container.winfo_children():
            widget.destroy()
        self.tab_buttons = {}

        for tab_name in self.tabs:
            if tab_name in ("Import Type", "AC Event Name", "AC Meeting Bodies"):
                active_color, hover_color = "#2e7d32", "#1b5e20"
            elif tab_name == "Checklist":
                active_color, hover_color = "#d4a500", "#b8900a"
            else:
                active_color, hover_color = "#1f538d", "#14375e"

            is_active = (tab_name == self.current_tab)
            btn = ctk.CTkButton(
                self.button_container,
                text=tab_name,
                width=140, height=40,
                font=ctk.CTkFont(size=14, weight="bold"),
                fg_color=active_color if is_active else "#3a3a3a",
                hover_color=hover_color,
                corner_radius=8,
                command=lambda t=tab_name: self._switch_tab(t),
            )
            btn.pack(side="left", padx=5)
            self.tab_buttons[tab_name] = btn

    def _switch_tab(self, tab_name: str):
        """Hide current tab frame and show the new one."""
        if self.current_tab in self.tab_frames:
            self.tab_frames[self.current_tab].pack_forget()

        for name, btn in self.tab_buttons.items():
            if name == tab_name:
                if name in ("Import Type", "AC Event Name", "AC Meeting Bodies"):
                    btn.configure(fg_color="#2e7d32", hover_color="#1b5e20")
                elif name == "Checklist":
                    btn.configure(fg_color="#d4a500", hover_color="#b8900a")
                else:
                    btn.configure(fg_color="#1f538d", hover_color="#14375e")
            else:
                btn.configure(fg_color="#3a3a3a", hover_color="#4a4a4a")

        self.tab_frames[tab_name].pack(fill="both", expand=True)
        self.current_tab = tab_name

    def _get_event_categories(self) -> List[str]:
        """Unique event categories from meeting bodies (for Subtypes dropdown)."""
        categories = {b.event_category.strip() for b in self.meeting_bodies if b.event_category.strip()}
        result = [""] + sorted(categories)
        return result if len(result) > 1 else [""]

    # ================================================================
    # TAB BUILDERS  (thin wrappers — delegate to tab modules)
    # ================================================================

    def _build_import_type_tab(self):
        build_import_type_tab(self)

    def _select_import_type(self, import_type: str):
        select_import_type(self, import_type)

    def _build_checklist_tab(self):
        build_checklist_tab(self)

    def _build_checklist_items(self):
        build_checklist_items(self)

    def _open_link(self, url: str, needs_siteid: bool):
        open_link(self, url, needs_siteid)

    def _on_siteid_changed(self, *args):
        on_siteid_changed(self)

    def _save_checklist_state(self):
        save_checklist_state(self)

    def _build_meeting_bodies_tab(self):
        build_meeting_bodies_tab(self)

    def _browse_base_dir(self):
        browse_base_dir(self)

    def _import_from_sheet(self):
        import_from_sheet(self)

    def _auto_scan_meetings(self):
        auto_scan_meetings(self)

    def _add_meeting_body(self):
        add_meeting_body(self)

    def _remove_selected_bodies(self):
        remove_selected_bodies(self)

    def _clear_all_meeting_bodies(self):
        clear_all_meeting_bodies(self)

    def _refresh_meeting_bodies_display(self):
        refresh_meeting_bodies_display(self)

    def _update_meeting_time(self, index: int):
        update_meeting_time(self, index)

    def _browse_body_folder(self, index: int):
        browse_body_folder(self, index)

    def _build_subtypes_tab(self):
        build_subtypes_tab(self)

    def _create_advanced_matching_section(self, parent):
        create_advanced_matching_section(self, parent)

    def _update_scan_filter_options(self):
        update_scan_filter_options(self)

    def _toggle_advanced_matching(self):
        toggle_advanced_matching(self)

    def _on_pattern_selected(self, choice):
        on_pattern_selected(self, choice)

    def _auto_scan_subtypes(self):
        auto_scan_subtypes(self)

    def _next_from_meeting_bodies(self):
        next_from_meeting_bodies(self)

    def _next_from_subtypes(self):
        next_from_subtypes(self)

    def _add_subtype(self):
        add_subtype(self)

    def _remove_selected_subtypes(self):
        remove_selected_subtypes(self)

    def _clear_all_subtypes(self):
        clear_all_subtypes(self)

    def _refresh_subtypes_display(self):
        refresh_subtypes_display(self)

    def _update_subtype_time(self, index: int):
        update_subtype_time(self, index)

    def _move_subtype(self, index: int, direction: int):
        move_subtype(self, index, direction)

    def _drag_release(self, event, index):
        drag_release(self, event, index)

    def _build_file_errors_tab(self):
        build_file_errors_tab(self)

    def _refresh_file_errors_display(self):
        refresh_file_errors_display(self)

    def _preview_selected_pdf(self):
        preview_selected_pdf(self)

    def _run_ocr_on_selected(self):
        run_ocr_on_selected(self)

    def _apply_ocr_suggestion(self, error_index):
        apply_ocr_suggestion(self, error_index)

    def _edit_ocr_suggestion(self, error_index):
        edit_ocr_suggestion(self, error_index)

    def _rename_selected_errors(self):
        rename_selected_errors(self)

    def _rename_single_file(self, error: FileError):
        rename_single_file(self, error)

    def _bulk_rename_files(self, errors):
        bulk_rename_files(self, errors)

    def _apply_bulk_rename(self, errors, pattern, text, dialog):
        apply_bulk_rename(self, errors, pattern, text, dialog)

    def _open_error_location(self):
        open_error_location(self)

    def _mark_errors_resolved(self):
        mark_errors_resolved(self)

    def _remove_duplicates(self):
        remove_duplicates(self)

    def _select_all_errors(self):
        select_all_errors(self)

    def _deselect_all_errors(self):
        deselect_all_errors(self)

    def _rerun_import(self):
        rerun_import(self)

    def _build_run_tab(self):
        build_run_tab(self)

    def _run_dry_run(self):
        run_dry_run(self)

    def _run_import(self):
        run_import(self)

    def _execute_import(self, dry_run=True):
        execute_import(self, dry_run)

    def _mock_import(self, config):
        mock_import(self, config)

    def _display_results(self, result: dict):
        display_results(self, result)

    def _populate_file_errors(self, errors: list, dry_run: bool = False):
        populate_file_errors(self, errors, dry_run)

    def _open_review_folders(self):
        open_review_folders(self)

    def _export_excel(self):
        export_excel(self)

    def _build_ac_event_name_tab(self):
        build_ac_event_name_tab(self)

    def _ac_browse_file(self):
        ac_browse_file(self)

    def _ac_populate_event_name_rows(self, flagged: list):
        ac_populate_event_name_rows(self, flagged)

    def _ac_next_to_meeting_bodies(self):
        ac_next_to_meeting_bodies(self)

    def _build_ac_meeting_bodies_tab(self):
        build_ac_meeting_bodies_tab(self)

    def _ac_preload_meeting_bodies(self):
        ac_preload_meeting_bodies(self)

    def _ac_save_time(self, category: str):
        ac_save_time(self, category)

    def _ac_run_import(self):
        ac_run_import(self)

    # ================================================================
    # MANAGERS
    # ================================================================

    def _initialize_managers(self):
        """Initialize SharePoint and backend managers in background thread."""
        try:
            loading = ctk.CTkToplevel(self)
            loading.title("Initializing")
            loading.geometry("400x150")
            loading.transient(self)
            loading.grab_set()

            label = ctk.CTkLabel(loading, text="Checking for updates...",
                                 font=ctk.CTkFont(size=14))
            label.pack(pady=30)
            progress = ctk.CTkProgressBar(loading, mode="indeterminate")
            progress.pack(pady=10, padx=40, fill="x")
            progress.start()

            def init_thread():
                try:
                    script_path = None
                    if SharePointManager:
                        self.sharepoint_manager = SharePointManager()
                        script_path = self.sharepoint_manager.get_latest_script()
                    if not script_path:
                        local_scripts = list(Path(".").glob("Meeting_Import_Stable_v*.py"))
                        if local_scripts:
                            local_scripts.sort(reverse=True)
                            script_path = str(local_scripts[0])
                            status_text = f"✓ Using local script: {local_scripts[0].name}"
                        else:
                            status_text = "⚠ No script available\nPlace backend script in this folder or sync OneDrive"
                    else:
                        status_text = f"✓ Script ready: {Path(script_path).name}"

                    if BackendWrapper:
                        self.backend_wrapper = BackendWrapper(script_path if script_path else None)

                    self.after(0, lambda: label.configure(text=status_text))
                    self.after(2000, loading.destroy)
                except Exception as e:
                    self.after(0, lambda: label.configure(
                        text=f"⚠ Initialization warning: {str(e)}\nPlace backend script in app folder"))
                    self.after(3000, loading.destroy)

            threading.Thread(target=init_thread, daemon=True).start()
        except Exception as e:
            messagebox.showwarning("Initialization", f"Started with warnings: {str(e)}")

    # ================================================================
    # CONFIGURATION PERSISTENCE
    # ================================================================

    def _save_config(self):
        """Save configuration to JSON in the base directory."""
        base_dir = self.base_dir.get().strip()
        if not base_dir:
            return
        config = {
            "base_dir": base_dir,
            "date_format": self.date_format.get(),
            "meeting_bodies": [b.to_dict() for b in self.meeting_bodies],
            "subtypes": [s.to_dict() for s in self.subtypes],
            "checklist_state": self.checklist_state,
        }
        try:
            with open(Path(base_dir) / "meeting_import_config.json", "w") as f:
                json.dump(config, f, indent=2)
        except Exception as e:
            print(f"Error saving config: {e}")

    def _load_config(self):
        """Load configuration from JSON in the base directory."""
        base_dir = self.base_dir.get().strip()
        if not base_dir:
            global_path = Path.home() / ".meeting_import_gui_last_basedir.json"
            if global_path.exists():
                try:
                    with open(global_path) as f:
                        base_dir = json.load(f).get("last_base_dir", "")
                    if base_dir:
                        self.base_dir.set(base_dir)
                except Exception:
                    pass
        if not base_dir:
            return

        config_path = Path(base_dir) / "meeting_import_config.json"
        if not config_path.exists():
            return
        try:
            with open(config_path) as f:
                config = json.load(f)
            self.base_dir.set(config.get("base_dir", base_dir))
            self.date_format.set(config.get("date_format", "US DATE"))
            self.meeting_bodies = [MeetingBody.from_dict(d) for d in config.get("meeting_bodies", [])]
            self.subtypes = [Subtype.from_dict(d) for d in config.get("subtypes", [])]
            self.checklist_state = config.get("checklist_state", {})
            self._refresh_meeting_bodies_display()
            self._refresh_subtypes_display()
            global_path = Path.home() / ".meeting_import_gui_last_basedir.json"
            try:
                with open(global_path, "w") as f:
                    json.dump({"last_base_dir": base_dir}, f)
            except Exception:
                pass
        except Exception as e:
            print(f"Error loading config: {e}")
