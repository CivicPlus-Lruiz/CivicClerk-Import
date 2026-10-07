"""
Desktop PDF Viewer using PyMuPDF (fitz) + tkinter
Displays only the first page of the PDF for fast, lightweight previewing.

Requirements:
    pip install pymupdf Pillow

Usage:
    python pdf_viewer.py
    python pdf_viewer.py path/to/file.pdf
"""

import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ImportError:
    print("PyMuPDF not found. Install it with: pip install pymupdf")
    sys.exit(1)

try:
    from PIL import Image, ImageTk
except ImportError:
    print("Pillow not found. Install it with: pip install Pillow")
    sys.exit(1)


class PDFViewer:
    def __init__(self, root: tk.Tk):
        self.root = root
        self._is_window = isinstance(root, (tk.Tk, tk.Toplevel))
        if self._is_window:
            self.root.title("PDF Viewer")
            self.root.geometry("900x700")
        self.root.configure(bg="#1e1e2e")

        self.doc: fitz.Document | None = None
        self.zoom = 1.0
        self.photo_image = None  # Keep reference to prevent GC

        self._build_ui()
        self._bind_keys()

    # ------------------------------------------------------------------ #
    #  UI Construction                                                     #
    # ------------------------------------------------------------------ #

    def _build_ui(self):
        """Build the full application UI."""
        # ── Top toolbar ──────────────────────────────────────────────────
        toolbar = tk.Frame(self.root, bg="#313244", pady=6)
        toolbar.pack(side=tk.TOP, fill=tk.X)

        btn_style = {"bg": "#45475a", "fg": "#cdd6f4", "relief": tk.FLAT,
                     "padx": 10, "pady": 4, "cursor": "hand2",
                     "activebackground": "#585b70", "activeforeground": "#cdd6f4"}

        tk.Button(toolbar, text="📂  Open", command=self.open_file, **btn_style).pack(side=tk.LEFT, padx=(8, 2))

        # Page indicator (read-only, first page only)
        self.page_label = tk.Label(toolbar, text="Page 1 (preview only)",
                                   bg="#313244", fg="#a6adc8")
        self.page_label.pack(side=tk.LEFT, padx=12)

        # Zoom controls
        tk.Button(toolbar, text="🔍−", command=self.zoom_out, **btn_style).pack(side=tk.LEFT, padx=2)
        self.zoom_label = tk.Label(toolbar, text="100%", bg="#313244", fg="#a6adc8", width=5)
        self.zoom_label.pack(side=tk.LEFT)
        tk.Button(toolbar, text="🔍+", command=self.zoom_in,  **btn_style).pack(side=tk.LEFT, padx=2)
        tk.Button(toolbar, text="⊡ Fit",  command=self.zoom_fit, **btn_style).pack(side=tk.LEFT, padx=(2, 8))

        # File name label (right-aligned)
        self.file_label = tk.Label(toolbar, text="No file open", bg="#313244",
                                   fg="#6c7086", font=("Courier", 9))
        self.file_label.pack(side=tk.RIGHT, padx=10)

        # ── Scrollable canvas area ────────────────────────────────────────
        canvas_frame = tk.Frame(self.root, bg="#1e1e2e")
        canvas_frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(canvas_frame, bg="#181825", highlightthickness=0,
                                cursor="fleur")
        self.v_scroll = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL,
                                      command=self.canvas.yview)
        self.h_scroll = ttk.Scrollbar(self.root, orient=tk.HORIZONTAL,
                                      command=self.canvas.xview)

        self.canvas.configure(yscrollcommand=self.v_scroll.set,
                              xscrollcommand=self.h_scroll.set)

        self.v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Mouse drag-to-scroll
        self.canvas.bind("<ButtonPress-1>",   self._drag_start)
        self.canvas.bind("<B1-Motion>",       self._drag_move)
        self.canvas.bind("<MouseWheel>",       self._mouse_wheel)      # Windows/macOS
        self.canvas.bind("<Button-4>",         self._mouse_wheel)      # Linux scroll up
        self.canvas.bind("<Button-5>",         self._mouse_wheel)      # Linux scroll down

        # ── Status bar ────────────────────────────────────────────────────
        self.status = tk.Label(self.root, text="Open a PDF to get started.",
                               bg="#313244", fg="#6c7086", anchor=tk.W, padx=10)
        self.status.pack(side=tk.BOTTOM, fill=tk.X)

    # ------------------------------------------------------------------ #
    #  Key bindings                                                        #
    # ------------------------------------------------------------------ #

    def _bind_keys(self):
        self.root.bind("<plus>",  lambda e: self.zoom_in())
        self.root.bind("<minus>", lambda e: self.zoom_out())
        self.root.bind("<o>",     lambda e: self.open_file())

    # ------------------------------------------------------------------ #
    #  File handling                                                       #
    # ------------------------------------------------------------------ #

    def open_file(self, path: str | None = None):
        if path is None:
            path = filedialog.askopenfilename(
                title="Open PDF",
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
            )
        if not path:
            return
        try:
            if self.doc:
                self.doc.close()
            self.doc = fitz.open(path)
            total_pages = len(self.doc)
            self.zoom = 1.0
            self.file_label.config(text=Path(path).name)
            if self._is_window:
                self.root.title(f"PDF Viewer — {Path(path).name}")
            self.page_label.config(
                text=f"Page 1 of {total_pages} (preview only)"
                if total_pages > 1 else "Page 1 of 1"
            )
            self._render_page()
        except Exception as exc:
            messagebox.showerror("Error", f"Could not open file:\n{exc}")

    # ------------------------------------------------------------------ #
    #  Rendering                                                           #
    # ------------------------------------------------------------------ #

    def _render_page(self):
        if not self.doc:
            return

        page = self.doc[0]  # Always render the first page only
        mat = fitz.Matrix(self.zoom, self.zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)

        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        self.photo_image = ImageTk.PhotoImage(img)

        self.canvas.delete("all")
        cw = self.canvas.winfo_width() or 900
        x = max(pix.width // 2, cw // 2)
        self.canvas.create_image(x, 10, anchor=tk.N, image=self.photo_image)
        self.canvas.configure(scrollregion=(0, 0,
                                             max(pix.width + 40, cw),
                                             pix.height + 20))
        self.canvas.yview_moveto(0)

        self.zoom_label.config(text=f"{int(self.zoom * 100)}%")
        self.status.config(
            text=f"Showing page 1 (first page preview)  |  "
                 f"Zoom: {int(self.zoom * 100)}%  |  "
                 f"Size: {pix.width} × {pix.height} px"
        )

    # ------------------------------------------------------------------ #
    #  Zoom                                                                #
    # ------------------------------------------------------------------ #

    def zoom_in(self):
        if self.zoom < 4.0:
            self.zoom = round(self.zoom + 0.25, 2)
            self._render_page()

    def zoom_out(self):
        if self.zoom > 0.25:
            self.zoom = round(self.zoom - 0.25, 2)
            self._render_page()

    def zoom_fit(self):
        """Fit page width to canvas width."""
        if not self.doc:
            return
        page = self.doc[0]
        cw = self.canvas.winfo_width() or 860
        self.zoom = round((cw - 40) / page.rect.width, 2)
        self._render_page()

    # ------------------------------------------------------------------ #
    #  Mouse interaction                                                   #
    # ------------------------------------------------------------------ #

    def _drag_start(self, event):
        self.canvas.scan_mark(event.x, event.y)

    def _drag_move(self, event):
        self.canvas.scan_dragto(event.x, event.y, gain=1)

    def _mouse_wheel(self, event):
        # Cross-platform scroll
        if event.num == 4 or event.delta > 0:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5 or event.delta < 0:
            self.canvas.yview_scroll(1, "units")


# ------------------------------------------------------------------ #
#  Entry point                                                         #
# ------------------------------------------------------------------ #

if __name__ == "__main__":
    root = tk.Tk()

    # Style the ttk scrollbars to match the dark theme
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Vertical.TScrollbar",   background="#45475a", troughcolor="#1e1e2e", bordercolor="#1e1e2e")
    style.configure("Horizontal.TScrollbar", background="#45475a", troughcolor="#1e1e2e", bordercolor="#1e1e2e")

    app = PDFViewer(root)

    # Optionally open a file passed as CLI argument
    if len(sys.argv) > 1:
        root.after(100, lambda: app.open_file(sys.argv[1]))

    root.mainloop()
