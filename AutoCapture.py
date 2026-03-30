"""
Screenshot → PDF Appender
=========================
Requirements (install via pip):
    pip install pillow pypdf reportlab pynput

Run:
    python screenshot_to_pdf.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import time
import os
import io

# ── lazy imports so the app shows a clear error if something is missing ──────
try:
    from PIL import Image, ImageGrab, ImageTk
except ImportError:
    raise SystemExit("Pillow not found.  Run:  pip install pillow")

try:
    import pypdf
except ImportError:
    raise SystemExit("pypdf not found.  Run:  pip install pypdf")

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas as rl_canvas
except ImportError:
    raise SystemExit("reportlab not found.  Run:  pip install reportlab")

try:
    from pynput import keyboard as pynput_keyboard
    from pynput import mouse as pynput_mouse
    PYNPUT_OK = True
except ImportError:
    PYNPUT_OK = False          # hotkey mode will warn the user


# ─────────────────────────────────────────────────────────────────────────────
#  Helpers
# ─────────────────────────────────────────────────────────────────────────────

def image_to_pdf_bytes(img: Image.Image) -> bytes:
    """Convert a PIL Image to a single-page PDF (bytes)."""
    buf = io.BytesIO()
    w, h = img.size
    c = rl_canvas.Canvas(buf, pagesize=(w, h))
    # save image to a temporary in-memory PNG so reportlab can read it
    img_buf = io.BytesIO()
    img.save(img_buf, format="PNG")
    img_buf.seek(0)
    from reportlab.lib.utils import ImageReader
    c.drawImage(ImageReader(img_buf), 0, 0, width=w, height=h)
    c.save()
    return buf.getvalue()


def append_image_to_pdf(pdf_path: str, img: Image.Image):
    """Append a PIL image as a new page to an existing (or new) PDF file."""
    new_page_bytes = image_to_pdf_bytes(img)
    new_page_reader = pypdf.PdfReader(io.BytesIO(new_page_bytes))

    writer = pypdf.PdfWriter()

    if os.path.exists(pdf_path):
        existing = pypdf.PdfReader(pdf_path)
        for page in existing.pages:
            writer.add_page(page)

    writer.add_page(new_page_reader.pages[0])

    with open(pdf_path, "wb") as f:
        writer.write(f)


# ─────────────────────────────────────────────────────────────────────────────
#  Area-selection overlay
# ─────────────────────────────────────────────────────────────────────────────

class AreaSelector:
    """
    Displays a full-screen semi-transparent overlay.
    User drags a rectangle; returns (x1, y1, x2, y2) or None on cancel.
    """

    def __init__(self, parent_root):
        self.result = None
        self._start = None
        self._rect_id = None

        self.top = tk.Toplevel(parent_root)
        self.top.attributes("-fullscreen", True)
        self.top.attributes("-alpha", 0.25)
        self.top.attributes("-topmost", True)
        self.top.configure(cursor="crosshair", bg="black")

        self.canvas = tk.Canvas(self.top, bg="black", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.canvas.bind("<ButtonPress-1>",   self._on_press)
        self.canvas.bind("<B1-Motion>",       self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        self.top.bind("<Escape>", lambda e: self._cancel())

        label = tk.Label(
            self.canvas,
            text="Drag to select region  •  Esc to cancel",
            bg="black", fg="white",
            font=("Segoe UI", 13, "bold"),
        )
        label.place(relx=0.5, rely=0.04, anchor="center")

    # ── events ────────────────────────────────────────────────────────────────
    def _on_press(self, event):
        self._start = (event.x, event.y)
        if self._rect_id:
            self.canvas.delete(self._rect_id)

    def _on_drag(self, event):
        if not self._start:
            return
        x0, y0 = self._start
        if self._rect_id:
            self.canvas.delete(self._rect_id)
        self._rect_id = self.canvas.create_rectangle(
            x0, y0, event.x, event.y,
            outline="#00ff99", width=2, dash=(4, 2)
        )

    def _on_release(self, event):
        if not self._start:
            return
        x0, y0 = self._start
        x1, y1 = event.x, event.y
        # normalise
        left   = min(x0, x1)
        top    = min(y0, y1)
        right  = max(x0, x1)
        bottom = max(y0, y1)
        if right - left > 4 and bottom - top > 4:
            self.result = (left, top, right, bottom)
        self.top.destroy()

    def _cancel(self):
        self.result = None
        self.top.destroy()

    def wait(self):
        self.top.wait_window()
        return self.result


# ─────────────────────────────────────────────────────────────────────────────
#  Clipboard preview popup
# ─────────────────────────────────────────────────────────────────────────────

class ClipboardPreview:
    """Show a popup with the clipboard image. Enter = accept, Esc = reject."""

    def __init__(self, parent, img: Image.Image):
        self.accepted = False
        self._img = img

        self.top = tk.Toplevel(parent)
        self.top.title("Clipboard capture – accept?")
        self.top.attributes("-topmost", True)
        self.top.resizable(False, False)

        # thumbnail
        thumb = img.copy()
        thumb.thumbnail((600, 400))
        self._tk_img = ImageTk.PhotoImage(thumb)
        tk.Label(self.top, image=self._tk_img).pack(padx=10, pady=10)

        bar = tk.Frame(self.top)
        bar.pack(pady=(0, 10))
        tk.Button(bar, text="✔  Accept  (Enter)", width=18,
                  command=self._accept, bg="#2ecc71", fg="white",
                  font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Button(bar, text="✘  Reject  (Esc)", width=18,
                  command=self._reject, bg="#e74c3c", fg="white",
                  font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=5)

        self.top.bind("<Return>", lambda e: self._accept())
        self.top.bind("<Escape>", lambda e: self._reject())
        self.top.focus_force()

    def _accept(self):
        self.accepted = True
        self.top.destroy()

    def _reject(self):
        self.accepted = False
        self.top.destroy()

    def wait(self):
        self.top.wait_window()
        return self.accepted


# ─────────────────────────────────────────────────────────────────────────────
#  Tooltip helper
# ─────────────────────────────────────────────────────────────────────────────

class Tooltip:
    def __init__(self, widget, text: str):
        self._widget = widget
        self._text   = text
        self._tip    = None
        widget.bind("<Enter>", self._show)
        widget.bind("<Leave>", self._hide)

    def _show(self, _event=None):
        x = self._widget.winfo_rootx() + 20
        y = self._widget.winfo_rooty() + self._widget.winfo_height() + 4
        self._tip = tk.Toplevel(self._widget)
        self._tip.wm_overrideredirect(True)
        self._tip.wm_geometry(f"+{x}+{y}")
        tk.Label(
            self._tip, text=self._text,
            background="#ffffcc", relief="solid", borderwidth=1,
            font=("Segoe UI", 9), wraplength=240, justify=tk.LEFT,
            padx=6, pady=4
        ).pack()

    def _hide(self, _event=None):
        if self._tip:
            self._tip.destroy()
            self._tip = None


# ─────────────────────────────────────────────────────────────────────────────
#  Main Application
# ─────────────────────────────────────────────────────────────────────────────

class App(tk.Tk):

    HOTKEY_DEFAULT = "f6"

    def __init__(self):
        super().__init__()
        self.title("AutoCapture")
        self.resizable(False, False)
        self.attributes("-topmost", False)

        # ── state ─────────────────────────────────────────────────────────────
        self._capture_region   = None   # (x1,y1,x2,y2)
        self._click_position   = None   # (x, y)
        self._hotkey_str       = self.HOTKEY_DEFAULT
        self._hotkey_listener  = None   # pynput listener
        self._clipboard_thread = None
        self._running          = False
        self._stop_event       = threading.Event()

        # tk variables
        self._autoscroll_on     = tk.BooleanVar(value=False)
        self._scroll_direction  = tk.StringVar(value="vertical")
        self._scroll_reverse    = tk.BooleanVar(value=False)
        self._scroll_pixels     = tk.IntVar(value=10)
        self._approx_loops      = tk.IntVar(value=10)

        # tk variables
        self._pdf_name     = tk.StringVar()
        self._pdf_location = tk.StringVar()
        self._autonext_on  = tk.BooleanVar(value=False)
        self._loop_count   = tk.IntVar(value=5)
        self._always_top   = tk.BooleanVar(value=False)
        self._upscale      = tk.BooleanVar(value=False)

        self._build_ui()
        self._always_top.trace_add("write", self._toggle_topmost)

    # ─────────────────────────────────────────────────────────────────────────
    #  UI construction
    # ─────────────────────────────────────────────────────────────────────────

    def _section(self, parent, row, text):
        """Render a bold visible section label spanning all columns."""
        tk.Label(parent, text=text,
                 font=("Segoe UI", 10, "bold"), fg="#2c3e50",
                 anchor="w").grid(row=row, column=0, columnspan=6,
                                  sticky="ew", padx=6, pady=(10, 2))
        ttk.Separator(parent, orient="horizontal").grid(
            row=row + 1, column=0, columnspan=6, sticky="ew", padx=4, pady=(0, 4))

    def _build_ui(self):
        PAD = dict(padx=6, pady=5)
        FONT_LABEL = ("Segoe UI", 9)
        FONT_BTN   = ("Segoe UI", 9, "bold")

        outer = tk.Frame(self, padx=14, pady=12)
        outer.pack(fill=tk.BOTH, expand=True)

        # ── App title ────────────────────────────────────────────────────────
        tk.Label(outer, text="⚡  AutoCapture",
                 font=("Segoe UI", 14, "bold"), fg="#1a252f").grid(
            row=0, column=0, columnspan=6, sticky="w", pady=(0, 4))

        # ── Section 0: IMPORTANT ─────────────────────────────────────────────
        self._section(outer, 2, "0.  IMPORTANT!")
        tk.Label(outer,
                 text="Disable animations before use — Windows 11: Settings → Accessibility → Animation Effects (Off)",
                 font=("Segoe UI", 8), fg="#e74c3c",
                 wraplength=520, justify=tk.LEFT).grid(
            row=4, column=0, columnspan=6, sticky="w", padx=6, pady=(0, 4))

        # ── Section 1: Physical Location ─────────────────────────────────────
        self._section(outer, 5, "1.  Physical Location")

        tk.Label(outer, text="PDF name:", font=FONT_LABEL).grid(
            row=7, column=0, sticky="e", **PAD)
        tk.Entry(outer, textvariable=self._pdf_name, width=18,
                 font=FONT_LABEL).grid(row=7, column=1, sticky="w", **PAD)
        tk.Label(outer, text="Location:", font=FONT_LABEL).grid(
            row=7, column=2, sticky="e", **PAD)
        tk.Entry(outer, textvariable=self._pdf_location, width=22,
                 font=FONT_LABEL).grid(row=7, column=3, sticky="w", **PAD)
        tk.Button(outer, text="Browse…", font=FONT_BTN,
                  command=self._browse).grid(row=7, column=4, **PAD)

        # ── Section 2: Automation Shortcut ───────────────────────────────────
        self._section(outer, 8, "2.  Automation Shortcut")

        self._hotkey_label = tk.Label(
            outer, text=f"Hotkey: [{self._hotkey_str.upper()}]",
            font=FONT_LABEL, fg="#555")
        self._hotkey_label.grid(row=10, column=0, sticky="e", **PAD)

        tk.Button(outer, text="Edit Hotkey", font=FONT_BTN,
                  command=self._edit_hotkey).grid(row=10, column=1, **PAD)
        tk.Button(outer, text="Edit Hotkey Area", font=FONT_BTN,
                  command=self._edit_hotkey_area).grid(row=10, column=2, **PAD)

        self._hotkey_area_label = tk.Label(outer, text="(no area set)",
                                           font=FONT_LABEL, fg="#888")
        self._hotkey_area_label.grid(row=10, column=3, sticky="w", **PAD)

        tip1 = tk.Label(outer, text="(?)", font=FONT_LABEL, fg="#3498db",
                        cursor="question_arrow")
        tip1.grid(row=10, column=4, **PAD)
        Tooltip(tip1, "Defines the screen region captured when you press the hotkey.")

        # ── Section 3a: AutoClickNext ─────────────────────────────────────────
        self._section(outer, 11, "3(a).  AutoClickNext")

        tk.Checkbutton(outer, text="Autonext?", variable=self._autonext_on,
                       font=FONT_BTN, command=self._toggle_autonext_ui).grid(
            row=13, column=0, columnspan=2, sticky="w", **PAD)

        tip2 = tk.Label(outer, text="(?)", font=FONT_LABEL, fg="#3498db",
                        cursor="question_arrow")
        tip2.grid(row=13, column=2, sticky="w", **PAD)
        Tooltip(tip2, "Automates a capture → click → capture loop N times.")

        self._autonext_btn = tk.Button(
            outer, text="Set Click Location", font=FONT_BTN,
            command=self._set_click_location, state=tk.DISABLED)
        self._autonext_btn.grid(row=13, column=3, **PAD)

        self._click_pos_label = tk.Label(outer, text="(none)", font=FONT_LABEL,
                                         fg="#888")
        self._click_pos_label.grid(row=13, column=4, sticky="w", **PAD)

        # loop count
        self._loop_frame = tk.Frame(outer)
        self._loop_frame.grid(row=13, column=5, sticky="w", **PAD)
        tk.Label(self._loop_frame, text="Loops:", font=FONT_LABEL).pack(
            side=tk.LEFT)
        self._loop_spin = tk.Spinbox(
            self._loop_frame, from_=1, to=999, width=5,
            textvariable=self._loop_count, font=FONT_LABEL,
            state=tk.DISABLED)
        self._loop_spin.pack(side=tk.LEFT, padx=(3, 0))

        # ── Section 3b: AutoScrollNext ────────────────────────────────────────
        self._section(outer, 14, "3(b).  AutoScrollNext")

        tk.Checkbutton(outer, text="Autoscroll?", variable=self._autoscroll_on,
                       font=FONT_BTN, command=self._toggle_autoscroll_ui).grid(
            row=16, column=0, columnspan=2, sticky="w", **PAD)

        tip_as = tk.Label(outer, text="(?)", font=FONT_LABEL, fg="#3498db",
                          cursor="question_arrow")
        tip_as.grid(row=16, column=2, sticky="w", **PAD)
        Tooltip(tip_as, (
            "Autoscroll mode: captures the hotkey area, scrolls by the exact pixel "
            "height (vertical) or width (horizontal) of that area, captures again, "
            "and repeats. Mutual-exclusive with Autonext."
        ))

        tk.Label(outer, text="Direction:", font=FONT_LABEL).grid(
            row=17, column=0, sticky="e", **PAD)
        rb_v = tk.Radiobutton(outer, text="Vertical (scroll ↓)",
                              variable=self._scroll_direction, value="vertical",
                              font=FONT_LABEL, state=tk.DISABLED)
        rb_v.grid(row=17, column=1, sticky="w", **PAD)
        rb_h = tk.Radiobutton(outer, text="Horizontal (scroll →)",
                              variable=self._scroll_direction, value="horizontal",
                              font=FONT_LABEL, state=tk.DISABLED)
        rb_h.grid(row=17, column=2, columnspan=2, sticky="w", **PAD)

        rb_rev = tk.Checkbutton(outer, text="Reverse scroll",
                                variable=self._scroll_reverse,
                                font=FONT_LABEL, state=tk.DISABLED)
        rb_rev.grid(row=17, column=4, sticky="w", **PAD)

        tk.Label(outer, text="Scroll units:", font=FONT_LABEL).grid(
            row=18, column=0, sticky="e", **PAD)
        spin_px = tk.Spinbox(outer, from_=1, to=9999, width=6,
                             textvariable=self._scroll_pixels,
                             font=FONT_LABEL, state=tk.DISABLED)
        spin_px.grid(row=18, column=1, sticky="w", **PAD)

        tip_px = tk.Label(outer, text="(?)", font=FONT_LABEL, fg="#3498db",
                          cursor="question_arrow")
        tip_px.grid(row=18, column=2, sticky="w", **PAD)
        Tooltip(tip_px, (
            "Scroll units per step (OS scroll ticks, NOT screen pixels).\n"
            "Start with a small number like 5, run one test capture, "
            "and adjust up/down until the two captures align perfectly with no overlap and no gap.\n"
            "Typical values: 3–15 depending on your OS scroll speed settings."
        ))

        tk.Label(outer, text="Approx loops:", font=FONT_LABEL).grid(
            row=18, column=3, sticky="e", **PAD)
        spin_loops = tk.Spinbox(outer, from_=1, to=9999, width=6,
                                textvariable=self._approx_loops,
                                font=FONT_LABEL, state=tk.DISABLED)
        spin_loops.grid(row=18, column=4, sticky="w", **PAD)

        tip_loops = tk.Label(outer, text="(?)", font=FONT_LABEL, fg="#3498db",
                             cursor="question_arrow")
        tip_loops.grid(row=18, column=5, sticky="w", **PAD)
        Tooltip(tip_loops, (
            "Approximate number of scroll-captures needed to cover the whole document.\n"
            "Example: a 57-page PDF viewed at 50% zoom needs ~114 loops.\n"
            "After this many captures a confirmation dialog appears after every "
            "additional capture until you press Stop."
        ))

        # collect widgets to enable/disable together
        self._autoscroll_widgets = [rb_v, rb_h, rb_rev, spin_px, spin_loops]

        # ── Section 4: Automate and Relax ─────────────────────────────────────
        self._section(outer, 19, "4.  Automate and Relax~")

        self._start_btn = tk.Button(
            outer, text="▶  Start Appending", font=FONT_BTN,
            bg="#27ae60", fg="white", width=18,
            command=self._start_appending)
        self._start_btn.grid(row=21, column=0, columnspan=2, **PAD)

        self._stop_btn = tk.Button(
            outer, text="■  Stop Appending", font=FONT_BTN,
            bg="#c0392b", fg="white", width=18,
            command=self._stop_appending, state=tk.DISABLED)
        self._stop_btn.grid(row=21, column=2, columnspan=2, **PAD)

        self._status_canvas = tk.Canvas(outer, width=16, height=16,
                                        highlightthickness=0)
        self._status_canvas.grid(row=21, column=4, **PAD)
        self._status_dot = self._status_canvas.create_oval(
            2, 2, 14, 14, fill="#888", outline="")

        self._status_label = tk.Label(outer, text="Stopped",
                                      font=FONT_LABEL, fg="#888")
        self._status_label.grid(row=21, column=5, sticky="w", **PAD)

        # ── Section: Utilities ────────────────────────────────────────────────
        self._section(outer, 22, "Utilities")

        tk.Checkbutton(outer, text="Stay Always On Top",
                       variable=self._always_top,
                       font=FONT_LABEL).grid(
            row=24, column=0, columnspan=3, sticky="w", **PAD)

        tk.Checkbutton(outer, text="Upscale image (2×) before appending",
                       variable=self._upscale,
                       font=FONT_LABEL).grid(
            row=24, column=3, columnspan=3, sticky="w", **PAD)

        self._info_bar = tk.Label(
            outer, text="Tip: Win+Shift+S captures to clipboard automatically.",
            font=("Segoe UI", 8), fg="#999", anchor="w")
        self._info_bar.grid(row=25, column=0, columnspan=6, sticky="ew",
                            pady=(4, 0))

    # ─────────────────────────────────────────────────────────────────────────
    #  UI callbacks
    # ─────────────────────────────────────────────────────────────────────────

    def _browse(self):
        folder = filedialog.askdirectory(title="Select PDF destination folder")
        if folder:
            self._pdf_location.set(folder)

    def _toggle_topmost(self, *_):
        self.attributes("-topmost", self._always_top.get())

    def _toggle_autonext_ui(self):
        if self._autonext_on.get():
            self._autoscroll_on.set(False)
            self._toggle_autoscroll_ui()
            self._autonext_btn.config(state=tk.NORMAL)
            self._loop_spin.config(state="normal")
        else:
            self._autonext_btn.config(state=tk.DISABLED)
            self._loop_spin.config(state=tk.DISABLED)

    def _toggle_autoscroll_ui(self):
        if self._autoscroll_on.get():
            self._autonext_on.set(False)
            self._toggle_autonext_ui_silent()
            for w in self._autoscroll_widgets:
                w.config(state=tk.NORMAL)
        else:
            for w in self._autoscroll_widgets:
                w.config(state=tk.DISABLED)

    def _toggle_autonext_ui_silent(self):
        """Disable autonext widgets without triggering mutual-exclusion loop."""
        self._autonext_btn.config(state=tk.DISABLED)
        self._loop_spin.config(state=tk.DISABLED)

    # ── Edit hotkey ────────────────────────────────────────────────────────

    def _edit_hotkey(self):
        if not PYNPUT_OK:
            messagebox.showwarning(
                "pynput missing",
                "Install pynput to enable hotkey capture:\n  pip install pynput")
            return

        win = tk.Toplevel(self)
        win.title("Press new hotkey")
        win.resizable(False, False)
        win.attributes("-topmost", True)
        win.grab_set()

        tk.Label(win, text="Press the key you want to use as hotkey.\n"
                           "(Esc = cancel)",
                 font=("Segoe UI", 10), padx=20, pady=16).pack()

        captured = [None]

        def on_press(key):
            try:
                name = key.char if hasattr(key, "char") and key.char else key.name
            except Exception:
                return
            if name == "esc":
                listener.stop()
                win.after(0, win.destroy)
                return
            captured[0] = name
            listener.stop()
            win.after(0, _finish)

        def _finish():
            if captured[0]:
                self._hotkey_str = captured[0]
                self._hotkey_label.config(
                    text=f"Hotkey: [{self._hotkey_str.upper()}]")
            win.destroy()

        listener = pynput_keyboard.Listener(on_press=on_press)
        listener.start()
        win.wait_window()

    # ── Edit hotkey area ───────────────────────────────────────────────────

    def _edit_hotkey_area(self):
        self.withdraw()
        self.after(300, self._run_area_selector, "hotkey")

    def _edit_autonext_area(self):
        self.withdraw()
        self.after(150, self._run_area_selector, "autonext")

    def _run_area_selector(self, mode):
        sel = AreaSelector(self)
        region = sel.wait()
        self.deiconify()
        if region:
            self._capture_region = region
            x1, y1, x2, y2 = region
            self._hotkey_area_label.config(
                text=f"({x1},{y1})→({x2},{y2})", fg="#2c3e50")
        else:
            self._set_status("Area selection cancelled.")

    # ── Set click location ─────────────────────────────────────────────────

    def _set_click_location(self):
        if not PYNPUT_OK:
            messagebox.showwarning(
                "pynput missing",
                "Install pynput to capture click location:\n  pip install pynput")
            return

        self._set_status("Click anywhere to set Autonext click position…")
        self.withdraw()
        self.after(300, self._listen_for_click)

    def _listen_for_click(self):
        captured = [None]

        def on_click(x, y, button, pressed):
            if pressed:
                captured[0] = (x, y)
                return False  # stop listener

        with pynput_mouse.Listener(on_click=on_click) as lst:
            lst.join()

        self.after(0, lambda: self._apply_click(captured[0]))

    def _apply_click(self, pos):
        self.deiconify()
        if pos:
            self._click_position = pos
            self._click_pos_label.config(
                text=f"({pos[0]}, {pos[1]})", fg="#2c3e50")
        self._set_status("Click position captured." if pos else "Cancelled.")

    # ─────────────────────────────────────────────────────────────────────────
    #  Start / Stop
    # ─────────────────────────────────────────────────────────────────────────

    def _validate(self) -> str | None:
        """Return an error string or None if valid."""
        if not self._pdf_name.get().strip():
            return "PDF name is required."
        if not self._pdf_location.get().strip():
            return "PDF destination folder is required."
        if not os.path.isdir(self._pdf_location.get().strip()):
            return "PDF destination is not a valid folder."
        if self._autonext_on.get():
            if self._capture_region is None:
                return "Set a capture area before starting Autonext."
            if self._click_position is None:
                return "Set a click location before starting Autonext."
            try:
                n = int(self._loop_spin.get())
                if n < 1:
                    raise ValueError
            except ValueError:
                return "Loop count must be a positive integer."
        return None

    def _pdf_path(self) -> str:
        name = self._pdf_name.get().strip()
        if not name.lower().endswith(".pdf"):
            name += ".pdf"
        return os.path.join(self._pdf_location.get().strip(), name)

    def _resolve_pdf_path(self) -> str | None:
        """
        If the target PDF already exists, ask the user what to do.
        Returns the final path to use, or None if the user cancelled.
        """
        path = self._pdf_path()
        if not os.path.exists(path):
            return path

        dlg = tk.Toplevel(self)
        dlg.title("PDF already exists")
        dlg.attributes("-topmost", True)
        dlg.resizable(False, False)
        dlg.grab_set()

        tk.Label(dlg,
                 text=f'"{os.path.basename(path)}" already exists.\nWhat would you like to do?',
                 font=("Segoe UI", 10), padx=20, pady=14).pack()

        choice = [None]

        def _append():
            choice[0] = "append"
            dlg.destroy()

        def _overwrite():
            choice[0] = "overwrite"
            dlg.destroy()

        def _rename():
            choice[0] = "rename"
            dlg.destroy()

        def _cancel():
            choice[0] = None
            dlg.destroy()

        bar = tk.Frame(dlg)
        bar.pack(pady=(0, 14))
        tk.Button(bar, text="Append to existing", width=20,
                  command=_append, bg="#2980b9", fg="white",
                  font=("Segoe UI", 9, "bold")).pack(pady=3)
        tk.Button(bar, text="Overwrite (new PDF, same name)", width=20,
                  command=_overwrite, bg="#e67e22", fg="white",
                  font=("Segoe UI", 9, "bold")).pack(pady=3)
        tk.Button(bar, text="Save as (name)(1).pdf", width=20,
                  command=_rename, bg="#27ae60", fg="white",
                  font=("Segoe UI", 9, "bold")).pack(pady=3)
        tk.Button(bar, text="Cancel", width=20,
                  command=_cancel,
                  font=("Segoe UI", 9)).pack(pady=3)

        dlg.bind("<Escape>", lambda e: _cancel())
        dlg.focus_force()
        dlg.wait_window()

        if choice[0] is None:
            return None
        if choice[0] == "append":
            return path
        if choice[0] == "overwrite":
            os.remove(path)
            return path
        # rename: find first free (name)(N).pdf
        base = self._pdf_name.get().strip()
        if base.lower().endswith(".pdf"):
            base = base[:-4]
        folder = self._pdf_location.get().strip()
        n = 1
        while True:
            candidate = os.path.join(folder, f"{base}({n}).pdf")
            if not os.path.exists(candidate):
                self._pdf_name.set(f"{base}({n})")
                return candidate
            n += 1

    def _start_appending(self):
        err = self._validate()
        if err:
            messagebox.showerror("Validation error", err)
            return

        final_path = self._resolve_pdf_path()
        if final_path is None:
            return   # user cancelled

        self._running = True
        self._stop_event.clear()
        self._start_btn.config(state=tk.DISABLED)
        self._stop_btn.config(state=tk.NORMAL)
        self._set_status("Running", running=True)

        if self._autonext_on.get():
            t = threading.Thread(target=self._autonext_loop, daemon=True)
            t.start()
        elif self._autoscroll_on.get():
            t = threading.Thread(target=self._autoscroll_loop, daemon=True)
            t.start()
        else:
            # start both clipboard listener and hotkey listener
            t_clip = threading.Thread(target=self._clipboard_loop, daemon=True)
            t_clip.start()
            if PYNPUT_OK:
                self._start_hotkey_listener()
            else:
                self._set_status("Running (clipboard only – pynput missing)")

    def _stop_appending(self):
        self._stop_event.set()
        self._running = False
        if self._hotkey_listener:
            try:
                self._hotkey_listener.stop()
            except Exception:
                pass
            self._hotkey_listener = None
        self._start_btn.config(state=tk.NORMAL)
        self._stop_btn.config(state=tk.DISABLED)
        self._set_status("Stopped")

    # ─────────────────────────────────────────────────────────────────────────
    #  Capture helpers
    # ─────────────────────────────────────────────────────────────────────────

    def _capture_region_image(self) -> Image.Image | None:
        if not self._capture_region:
            self._set_status("No capture area defined.")
            return None
        x1, y1, x2, y2 = self._capture_region
        self.withdraw()
        self.update_idletasks()
        time.sleep(0.1)  # wait for the app to hide properly
        img = ImageGrab.grab(bbox=(x1, y1, x2, y2))
        self.deiconify()
        return img

    def _save_capture(self, img: Image.Image):
        try:
            if self._upscale.get():
                img = img.resize((img.width * 2, img.height * 2), Image.LANCZOS)
            append_image_to_pdf(self._pdf_path(), img)
            self._set_status(f"Page appended → {os.path.basename(self._pdf_path())}")
        except Exception as exc:
            self._set_status(f"Error saving: {exc}")

    # ─────────────────────────────────────────────────────────────────────────
    #  Clipboard listener (polls every 0.8 s)
    # ─────────────────────────────────────────────────────────────────────────

    def _clipboard_loop(self):
        last_id = None
        while not self._stop_event.is_set():
            try:
                img = ImageGrab.grabclipboard()
                if isinstance(img, Image.Image):
                    img_id = hash(img.tobytes())
                    if img_id != last_id:
                        last_id = img_id
                        # show preview on main thread
                        self.after(0, self._show_clipboard_preview, img)
            except Exception:
                pass
            time.sleep(0.8)

    def _show_clipboard_preview(self, img: Image.Image):
        if not self._running:
            return
        prev = ClipboardPreview(self, img)
        accepted = prev.wait()
        if accepted:
            self._save_capture(img)

    # ─────────────────────────────────────────────────────────────────────────
    #  Hotkey listener
    # ─────────────────────────────────────────────────────────────────────────

    def _start_hotkey_listener(self):
        if not PYNPUT_OK:
            return
        hotkey_str = self._hotkey_str

        def on_press(key):
            if not self._running:
                return
            try:
                name = key.char if hasattr(key, "char") and key.char else key.name
            except Exception:
                return
            if name == hotkey_str:
                img = self._capture_region_image()
                if img:
                    self._save_capture(img)

        self._hotkey_listener = pynput_keyboard.Listener(on_press=on_press)
        self._hotkey_listener.daemon = True
        self._hotkey_listener.start()

    # ─────────────────────────────────────────────────────────────────────────
    #  Autonext loop
    # ─────────────────────────────────────────────────────────────────────────

    def _autonext_loop(self):
        if not PYNPUT_OK:
            self.after(0, messagebox.showwarning, "pynput missing",
                       "Autonext requires pynput:\n  pip install pynput")
            self.after(0, self._stop_appending)
            return

        n     = int(self._loop_spin.get())
        cx, cy = self._click_position
        mc    = pynput_mouse.Controller()

        for i in range(n):
            if self._stop_event.is_set():
                break
            # capture
            img = self._capture_region_image()
            if img:
                self._save_capture(img)
            # click (N-1 times — not after the last capture)
            if i < n - 1:
                time.sleep(0.3)
                mc.position = (cx, cy)
                mc.press(pynput_mouse.Button.left)
                time.sleep(0.05)
                mc.release(pynput_mouse.Button.left)
                time.sleep(0.5)

        self.after(0, self._stop_appending)
        self.after(0, self._set_status,
                   f"Autonext complete – {n} captures saved.")

    # ─────────────────────────────────────────────────────────────────────────
    #  Autoscroll loop
    # ─────────────────────────────────────────────────────────────────────────

    def _autoscroll_loop(self):
        if not PYNPUT_OK:
            self.after(0, messagebox.showwarning, "pynput missing",
                       "Autoscroll requires pynput:\n  pip install pynput")
            self.after(0, self._stop_appending)
            return

        if not self._capture_region:
            self.after(0, messagebox.showerror, "No capture area",
                       "Set a hotkey capture area before starting Autoscroll.")
            self.after(0, self._stop_appending)
            return

        x1, y1, x2, y2 = self._capture_region
        region_w = x2 - x1
        region_h = y2 - y1
        cx = (x1 + x2) // 2
        cy = (y1 + y2) // 2

        direction   = self._scroll_direction.get()   # "vertical" or "horizontal"
        reverse     = self._scroll_reverse.get()
        px_setting  = self._scroll_pixels.get()
        approx      = self._approx_loops.get()

        # pixel distance per scroll step
        if px_setting > 0:
            step_px = px_setting
        else:
            step_px = region_h if direction == "vertical" else region_w

        scroll_sign = -1 if not reverse else 1   # -1 = forward (down/right)
        mc = pynput_mouse.Controller()
        loop_count = 0

        while not self._stop_event.is_set():
            # --- capture ---
            img = self._capture_region_image()
            if img:
                self._save_capture(img)
            loop_count += 1

            if self._stop_event.is_set():
                break

            # --- confirmation after approx loops ---
            if loop_count >= approx:
                confirmed = [None]
                event = threading.Event()

                def _ask():
                    dlg = tk.Toplevel(self)
                    dlg.title("Continue scrolling?")
                    dlg.attributes("-topmost", True)
                    dlg.resizable(False, False)
                    dlg.grab_set()
                    tk.Label(dlg,
                             text=(f"Reached {loop_count} loops "
                                   f"(approx target was {approx}).\n\n"
                                   "Continue capturing?\n\n"
                                   "  Enter = Yes, one more then ask again\n"
                                   "  Esc   = Stop appending now"),
                             font=("Segoe UI", 10), padx=20, pady=14,
                             justify=tk.LEFT).pack()
                    bar = tk.Frame(dlg)
                    bar.pack(pady=(0, 12))

                    def _yes():
                        confirmed[0] = "yes"
                        dlg.destroy()
                        event.set()

                    def _no():
                        confirmed[0] = "no"
                        dlg.destroy()
                        event.set()

                    tk.Button(bar, text="✔  Yes  (Enter)", width=16,
                              command=_yes, bg="#2ecc71", fg="white",
                              font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=6)
                    tk.Button(bar, text="✘  No  (Esc)", width=16,
                              command=_no, bg="#e74c3c", fg="white",
                              font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=6)
                    dlg.bind("<Return>", lambda e: _yes())
                    dlg.bind("<Escape>", lambda e: _no())
                    dlg.focus_force()

                self.after(0, _ask)
                event.wait()

                if confirmed[0] != "yes":
                    self.after(0, self._stop_appending)
                    return

            # --- scroll (units are OS scroll ticks, calibrate with Scroll px field) ---
            time.sleep(0.15)
            mc.position = (cx, cy)
            time.sleep(0.05)
            units = max(1, step_px) * scroll_sign
            if direction == "vertical":
                mc.scroll(0, units)          # negative = down
            else:
                mc.scroll(units, 0)          # positive = right
            time.sleep(0.4)

        self.after(0, self._stop_appending)
        self.after(0, self._set_status,
                   f"Autoscroll complete – {loop_count} captures saved.")

    # ─────────────────────────────────────────────────────────────────────────
    #  Status indicator
    # ─────────────────────────────────────────────────────────────────────────

    def _set_status(self, text: str, running: bool = False):
        self._status_label.config(text=text, fg="#27ae60" if running else "#888")
        if running:
            self._blink(True)
        else:
            self._status_canvas.itemconfig(self._status_dot, fill="#888")

    def _blink(self, state: bool):
        if not self._running:
            self._status_canvas.itemconfig(self._status_dot, fill="#888")
            return
        color = "#27ae60" if state else "#a8e6c0"
        self._status_canvas.itemconfig(self._status_dot, fill=color)
        self.after(600, self._blink, not state)


# ─────────────────────────────────────────────────────────────────────────────
#  Entry point
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()