"""
Import Type selector tab
"""

import customtkinter as ctk
from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from gui.meeting_import_app import MeetingImportGUI

# ================================================================
# TAB 0: IMPORT TYPE SELECTOR
# ================================================================

def build_import_type_tab(app):
    """Build the Import Type landing/selector tab."""
    tab = app.tab_frames["Import Type"]

    # Outer centering frame
    center = ctk.CTkFrame(tab, fg_color="transparent")
    center.place(relx=0.5, rely=0.5, anchor="center")

    # Title
    ctk.CTkLabel(
        center,
        text="Meeting Import Tool",
        font=ctk.CTkFont(size=32, weight="bold"),
    ).pack(pady=(0, 8))

    ctk.CTkLabel(
        center,
        text="Select your import type to get started",
        font=ctk.CTkFont(size=16),
        text_color="#aaaaaa",
    ).pack(pady=(0, 48))

    # Card row
    card_row = ctk.CTkFrame(center, fg_color="transparent")
    card_row.pack()

    # ── Standard Import card ──────────────────────────────────────
    std_card = ctk.CTkFrame(card_row, width=300, height=360,
                            fg_color="#1e2a3a", corner_radius=16,
                            border_width=2, border_color="#1f538d")
    std_card.pack(side="left", padx=18)
    std_card.pack_propagate(False)

    ctk.CTkLabel(std_card, text="📁", font=ctk.CTkFont(size=48)).pack(pady=(28, 8))
    ctk.CTkLabel(std_card, text="Standard Import",
                 font=ctk.CTkFont(size=20, weight="bold")).pack()
    ctk.CTkLabel(
        std_card,
        text="Import meeting files\nfrom a local folder structure.\nIncludes Subtypes configuration.",
        font=ctk.CTkFont(size=13),
        text_color="#aaaaaa",
        justify="center",
    ).pack(pady=14, padx=20)

    ctk.CTkButton(
        std_card,
        text="Select →",
        width=180,
        height=44,
        font=ctk.CTkFont(size=15, weight="bold"),
        fg_color="#1f538d",
        hover_color="#14375e",
        corner_radius=10,
        command=lambda: app._select_import_type("standard"),
    ).pack(pady=(8, 0))

    # ── Agenda Center Import card ─────────────────────────────────
    ac_card = ctk.CTkFrame(card_row, width=300, height=360,
                           fg_color="#1a2e1a", corner_radius=16,
                           border_width=2, border_color="#2e7d32")
    ac_card.pack(side="left", padx=18)
    ac_card.pack_propagate(False)

    ctk.CTkLabel(ac_card, text="📋", font=ctk.CTkFont(size=48)).pack(pady=(28, 8))
    ctk.CTkLabel(ac_card, text="Agenda Center Import",
                 font=ctk.CTkFont(size=20, weight="bold")).pack()
    ctk.CTkLabel(
        ac_card,
        text="Import from an Agenda Center\nexport. Streamlined workflow\nwithout Subtypes.",
        font=ctk.CTkFont(size=13),
        text_color="#aaaaaa",
        justify="center",
    ).pack(pady=14, padx=20)

    ctk.CTkButton(
        ac_card,
        text="Select →",
        width=180,
        height=44,
        font=ctk.CTkFont(size=15, weight="bold"),
        fg_color="#2e7d32",
        hover_color="#1b5e20",
        corner_radius=10,
        command=lambda: app._select_import_type("agenda_center"),
    ).pack(pady=(8, 0))

    # ── Excel Import card (NEW) ───────────────────────────────────
    xl_card = ctk.CTkFrame(card_row, width=300, height=360,
                           fg_color="#2a1e30", corner_radius=16,
                           border_width=2, border_color="#7b3fa0")
    xl_card.pack(side="left", padx=18)
    xl_card.pack_propagate(False)

    ctk.CTkLabel(xl_card, text="📊", font=ctk.CTkFont(size=48)).pack(pady=(28, 8))
    ctk.CTkLabel(xl_card, text="FTP/Azure Import",
                 font=ctk.CTkFont(size=20, weight="bold")).pack()
    ctk.CTkLabel(
        xl_card,
        text="Import from a multi-sheet\nExcel file or scan directly\nfrom FTP / Azure Storage.",
        font=ctk.CTkFont(size=13),
        text_color="#aaaaaa",
        justify="center",
    ).pack(pady=14, padx=20)

    ctk.CTkButton(
        xl_card,
        text="Select →",
        width=180,
        height=44,
        font=ctk.CTkFont(size=15, weight="bold"),
        fg_color="#7b3fa0",
        hover_color="#5c2e78",
        corner_radius=10,
        command=lambda: app._select_import_type("excel_import"),
    ).pack(pady=(8, 0))

    # Current selection badge (hidden until chosen)
    app._import_type_badge = ctk.CTkLabel(
        center, text="",
        font=ctk.CTkFont(size=14),
        text_color="#aaaaaa",
    )
    app._import_type_badge.pack(pady=(32, 0))


def select_import_type(app, import_type: str):
    """Handle import type selection and reveal relevant tabs."""
    app.import_type = import_type

    # Update the tab list to match selected type
    app.tabs = app.TABS_BY_TYPE[import_type]

    # Stay on Import Type tab but rebuild buttons
    app._rebuild_tab_buttons()

    # Update badge text
    labels = {
        "standard":      ("Standard Import",        "#1f538d"),
        "agenda_center": ("Agenda Center Import",    "#2e7d32"),
        "excel_import":  ("FTP/Azure Import",            "#7b3fa0"),
    }
    label, color = labels.get(import_type, ("Unknown", "#888888"))
    app._import_type_badge.configure(
        text=f"\u2713 {label} selected  —  use the tabs above to continue",
        text_color=color,
    )
