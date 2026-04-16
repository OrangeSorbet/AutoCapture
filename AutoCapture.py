import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import time
import os, sys
import io
import json
import cv2
import numpy as np
from PIL import Image
Image.MAX_IMAGE_PIXELS = None

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
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

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

def get_autocapture_path():
    import os, sys

    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)  # folder where EXE is
    else:
        base = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base, "autosave.autocapture")

def save_progress(app, path=None):
    if path is None:
        path = get_autocapture_path()

    data = {
        "pdf_name": app._pdf_name.get(),
        "pdf_location": app._pdf_location.get(),
        "capture_region": app._capture_region,
        "click_position": app._click_position,
        "autoscroll": app._autoscroll_on.get(),
        "autonext": app._autonext_on.get(),
        "scroll_direction": app._scroll_direction.get(),
        "scroll_reverse": app._scroll_reverse.get(),
        "scroll_pixels": app._scroll_pixels.get(),
        "approx_loops": app._approx_loops.get(),
        "loop_count": app._loop_count.get(),
        "always_top": app._always_top.get(),
        "upscale": app._upscale.get(),
        "autostitch": app._autostitch.get(),
    }

    with open(path, "w") as f:
        json.dump(data, f)


def load_progress(app, path):
    with open(path, "r") as f:
        data = json.load(f)

    app._pdf_name.set(data.get("pdf_name", ""))
    app._pdf_location.set(data.get("pdf_location", ""))

    app._capture_region = data.get("capture_region")
    app._click_position = data.get("click_position")

    app._autoscroll_on.set(data.get("autoscroll", False))
    app._autonext_on.set(data.get("autonext", False))
    app._scroll_direction.set(data.get("scroll_direction", "vertical"))
    app._scroll_reverse.set(data.get("scroll_reverse", False))
    app._scroll_pixels.set(data.get("scroll_pixels", 700))
    app._approx_loops.set(data.get("approx_loops", 5))
    app._loop_count.set(data.get("loop_count", 5))
    app._always_top.set(data.get("always_top", False))
    app._upscale.set(data.get("upscale", False))
    app._autostitch.set(data.get("autostitch", False))

# ─────────────────────────────────────────────────────────────────────────────
#  Main Application
# ─────────────────────────────────────────────────────────────────────────────

class App(tk.Tk):

    HOTKEY_DEFAULT = "f6"

    def __init__(self):
        super().__init__()
        self.title("AutoCapture")
        self.iconbitmap(resource_path("icon.ico"))
        # Get screen size
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()

        # Set window size as % of screen
        win_w = int(screen_w * 0.6)
        win_h = int(screen_h * 0.85)

        # Center the window
        pos_x = (screen_w - win_w) // 2
        pos_y = (screen_h - win_h) // 2

        self.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")

        # Optional: minimum size
        self.minsize(600, 500)
        self.resizable(True, True)
        self.attributes("-topmost", False)

        # ── state ─────────────────────────────────────────────────────────────
        self._capture_region   = None   # (x1,y1,x2,y2)
        self._click_position   = None   # (x, y)
        self._hotkey_str       = self.HOTKEY_DEFAULT
        self._hotkey_listener  = None   # pynput listener
        self._clipboard_thread = None
        self._running          = False
        self._stop_event       = threading.Event()
        self._stitch_frames    = []

        # tk variables
        self._autoscroll_on     = tk.BooleanVar(value=False)
        self._scroll_direction  = tk.StringVar(value="vertical")
        self._scroll_reverse    = tk.BooleanVar(value=False)
        self._scroll_pixels     = tk.DoubleVar(value=700.0)
        self._approx_loops      = tk.IntVar(value=5)

        # tk variables
        self._pdf_name     = tk.StringVar()
        self._pdf_location = tk.StringVar()
        self._autonext_on  = tk.BooleanVar(value=False)
        self._loop_count   = tk.IntVar(value=5)
        self._always_top   = tk.BooleanVar(value=False)
        self._upscale      = tk.BooleanVar(value=False)
        self._autostitch   = tk.BooleanVar(value=False)

        self._build_ui()
        # ── AUTOSAVE SETUP ──
        self._autosave_job = None
        self._AUTOSAVE_DELAY = 2000  # ms

        vars_to_watch = [
            self._pdf_name,
            self._pdf_location,
            self._autoscroll_on,
            self._scroll_direction,
            self._scroll_reverse,
            self._scroll_pixels,
            self._approx_loops,
            self._autonext_on,
            self._loop_count,
            self._always_top,
            self._upscale,
        ]

        for var in vars_to_watch:
            var.trace_add("write", self._schedule_autosave)
        self._always_top.trace_add("write", self._toggle_topmost)
        self.bind("<Escape>", lambda e: self._stop_appending() if self._running else None)
        if PYNPUT_OK:
            self._global_esc_listener = pynput_keyboard.Listener(
                on_press=self._on_global_key)
            self._global_esc_listener.daemon = True
            self._global_esc_listener.start()
            # ── AUTOLOAD ON START ──
        try:
            path = get_autocapture_path()
            if os.path.exists(path):
                load_progress(self, path)

                # update UI labels manually (important)
                if self._capture_region:
                    x1, y1, x2, y2 = self._capture_region
                    self._hotkey_area_label.config(
                        text=f"({x1},{y1})→({x2},{y2})", fg="#2c3e50"
                    )

                if self._click_position:
                    x, y = self._click_position
                    self._click_pos_label.config(
                        text=f"({x}, {y})", fg="#2c3e50"
                    )

                self._set_status("Loaded previous session.", color="#27ae60")
        except Exception as e:
            self._set_status(f"Load failed", color="#888")

    def _on_global_key(self, key):
        try:
            name = key.name if hasattr(key, "name") else None
        except Exception:
            return
        if name == "esc" and self._running:
            self.after(0, self._stop_appending)
            self.after(0, self.deiconify)

    # ─────────────────────────────────────────────────────────────────────────
    #  UI construction
    # ─────────────────────────────────────────────────────────────────────────

    def _section(self, parent, text, row=None):
        frame = tk.Frame(parent)
        frame.pack(fill="x", pady=(10, 5))

        tk.Label(frame, text=text,
                font=("Segoe UI", 10, "bold"),
                fg="#2c3e50", anchor="w").pack(anchor="w")

        ttk.Separator(frame, orient="horizontal").pack(fill="x", pady=(2, 6))

        content = tk.Frame(parent)
        content.pack(fill="x", padx=6)

        return content

    def _build_ui(self):
        PAD = dict(padx=6, pady=5)
        FONT_LABEL = ("Segoe UI", 9)
        FONT_BTN   = ("Segoe UI", 9, "bold")

        container = tk.Frame(self)
        container.pack(fill="both", expand=True)

        # ── App title ────────────────────────────────────────────────────────
        tk.Label(container, text="⚡  AutoCapture",
                 font=("Segoe UI", 14, "bold"), fg="#1a252f").pack(
            anchor="w", padx=6, pady=(0, 4))

        # ── Section 0: IMPORTANT ─────────────────────────────────────────────
        label = tk.Label(container,
            text="• Disable animations before use — Windows 11: Settings → Accessibility → Animation Effects (Off)\n"
                "• Press Esc at any time while appending to stop capturing without closing the app.\n"
                "• If the app doesn't respond, try pressing Esc again or restarting it.\n"
                "• For best results, use a consistent zoom level in your PDF viewer and test-capture once to adjust the scroll pixels.\n"
                "• If the app appears frozen after AutoStitch on 50+ pages, it is processing the stitch — please be patient.\n"
                "• Please press 'esc' only once if you want to stop the app, and wait a moment for it to respond. Repeatedly pressing 'esc' may cause issues.\n"
                "• If any field is unclear, hover over the (?) icons for detailed explanations and tips.\n"
                "• If any field is disabled, try enabling/disabling related checkboxes.\n"
                "• If results are undesirable, try capture again with same settings.\n"
                "• All settings are saved automatically and loaded on next start.",
            font=("Segoe UI", 13),
            fg="#e74c3c",
            wraplength=1,
            justify=tk.LEFT,
            anchor="w"
        )

        label.pack(fill="x", padx=4, pady=(0, 2))
        label.bind("<Configure>", lambda e: e.widget.config(wraplength=e.width - 10))

        # ── Section 1: Physical Location ─────────────────────────────────────
        self._section(container, "1.  Physical Location")
        row_frame = tk.Frame(container)
        row_frame.pack(fill="x", padx=6)
        tk.Label(row_frame, text="PDF name:", font=FONT_LABEL).pack(
            side="left", **PAD)
        tk.Entry(row_frame, textvariable=self._pdf_name, width=18,
                 font=FONT_LABEL).pack(side="left", **PAD)
        tk.Label(row_frame, text="Location:", font=FONT_LABEL).pack(
            side="left", **PAD)
        tk.Entry(row_frame, textvariable=self._pdf_location, width=22,
                 font=FONT_LABEL).pack(side="left", **PAD)
        tk.Button(row_frame, text="Browse…", font=FONT_BTN,
                  command=self._browse).pack(side="left", **PAD)

        # ── Section 2: Automation Shortcut ───────────────────────────────────
        self._section(container, "2.  Automation Shortcut")
        row_frame = tk.Frame(container)
        row_frame.pack(fill="x", padx=6)
        self._hotkey_label = tk.Label(
            row_frame, text=f"Hotkey: [{self._hotkey_str.upper()}]",
            font=FONT_LABEL, fg="#555")
        self._hotkey_label.pack(side="left", **PAD)
        tk.Button(row_frame, text="Edit Hotkey", font=FONT_BTN,
                  command=self._edit_hotkey).pack(side="left", **PAD)
        tk.Button(row_frame, text="Edit Hotkey Area", font=FONT_BTN,
                  command=self._edit_hotkey_area).pack(side="left", **PAD)
        self._hotkey_area_label = tk.Label(row_frame, text="(no area set)",
                                           font=FONT_LABEL, fg="#888")
        self._hotkey_area_label.pack(side="left", **PAD)
        self._peek_btn = tk.Button(row_frame, text="👁 Peek", font=FONT_LABEL)
        self._peek_btn.pack(side="left", **PAD)
        self._peek_btn.bind("<ButtonPress-1>", self._peek_press)
        self._peek_btn.bind("<ButtonRelease-1>", self._peek_release)
        tip1 = tk.Label(row_frame, text="(?)", font=FONT_LABEL, fg="#3498db",
                        cursor="question_arrow")
        tip1.pack(side="left", **PAD)
        Tooltip(tip1, "Defines the screen region captured when you press the hotkey.")

        # ── Section 3a: AutoClickNext ─────────────────────────────────────────
        self._section(container, "3(a).  AutoClickNext")
        row_frame = tk.Frame(container)
        row_frame.pack(fill="x", padx=6)
        tk.Checkbutton(row_frame, text="Autonext?", variable=self._autonext_on,
                       font=FONT_BTN, command=self._toggle_autonext_ui).pack(
            side="left", **PAD)
        tip2 = tk.Label(row_frame, text="(?)", font=FONT_LABEL, fg="#3498db",
                        cursor="question_arrow")
        tip2.pack(side="left", **PAD)
        Tooltip(tip2, "Automates a capture → click → capture loop N times.")
        self._autonext_btn = tk.Button(
            row_frame, text="Set Click Location", font=FONT_BTN,
            command=self._set_click_location, state=tk.DISABLED)
        self._autonext_btn.pack(side="left", **PAD)
        self._click_pos_label = tk.Label(row_frame, text="(none)", font=FONT_LABEL,
                                         fg="#888")
        self._click_pos_label.pack(side="left", **PAD)
        self._loop_frame = tk.Frame(row_frame)
        self._loop_frame.pack(side="left", **PAD)
        tk.Label(self._loop_frame, text="Loops:", font=FONT_LABEL).pack(
            side="left")
        self._loop_spin = tk.Spinbox(
            self._loop_frame, from_=1, to=999, width=5,
            textvariable=self._loop_count, font=FONT_LABEL,
            state=tk.DISABLED)
        self._loop_spin.pack(side="left", padx=(3, 0))

        # ── Section 3b: AutoScrollNext ────────────────────────────────────────
        self._section(container, "3(b).  AutoScrollNext")
        row_frame = tk.Frame(container)
        row_frame.pack(fill="x", padx=6)
        tk.Checkbutton(row_frame, text="Autoscroll?", variable=self._autoscroll_on,
                       font=FONT_BTN, command=self._toggle_autoscroll_ui).pack(
            side="left", **PAD)
        tip_as = tk.Label(row_frame, text="(?)", font=FONT_LABEL, fg="#3498db",
                          cursor="question_arrow")
        tip_as.pack(side="left", **PAD)
        Tooltip(tip_as, (
            "Autoscroll mode: captures the hotkey area, scrolls by the set number of "
            "OS scroll units, captures again, and repeats. Mutual-exclusive with Autonext."
        ))
        tk.Checkbutton(row_frame, text="AutoStitch (1-page seamless)",
                       variable=self._autostitch,
                       font=FONT_LABEL).pack(side="left", **PAD)
        tip_stitch = tk.Label(row_frame, text="(?)", font=FONT_LABEL, fg="#3498db",
                              cursor="question_arrow")
        tip_stitch.pack(side="left", **PAD)
        Tooltip(tip_stitch, (
            "AutoStitch buffers all captures and stitches them into a single seamless page on Stop.\n"
            "Recommended: 500 scroll units at 75% zoom in browser PDF viewer → ~3 captures per page for perfect overlap.\n"
            "No pages are written to the PDF until you press Stop."
        ))

        row_frame = tk.Frame(container)
        row_frame.pack(fill="x", padx=6)
        tk.Label(row_frame, text="Direction:", font=FONT_LABEL).pack(
            side="left", **PAD)
        rb_v = tk.Radiobutton(row_frame, text="Vertical (scroll ↓)",
                              variable=self._scroll_direction, value="vertical",
                              font=FONT_LABEL, state=tk.DISABLED)
        rb_v.pack(side="left", **PAD)
        rb_h = tk.Radiobutton(row_frame, text="Horizontal (scroll →)",
                              variable=self._scroll_direction, value="horizontal",
                              font=FONT_LABEL, state=tk.DISABLED)
        rb_h.pack(side="left", **PAD)
        rb_rev = tk.Checkbutton(row_frame, text="Reverse scroll",
                                variable=self._scroll_reverse,
                                font=FONT_LABEL, state=tk.DISABLED)
        rb_rev.pack(side="left", **PAD)

        row_frame = tk.Frame(container)
        row_frame.pack(fill="x", padx=6)
        tk.Label(row_frame, text="Scroll units:", font=FONT_LABEL).pack(
            side="left", **PAD)
        spin_px = tk.Spinbox(row_frame, from_=0.1, to=9999, increment=0.1, width=6,
                             textvariable=self._scroll_pixels, format="%.1f",
                             font=FONT_LABEL, state=tk.DISABLED)
        spin_px.pack(side="left", **PAD)
        tip_px = tk.Label(row_frame, text="(?)", font=FONT_LABEL, fg="#3498db",
                          cursor="question_arrow")
        tip_px.pack(side="left", **PAD)
        Tooltip(tip_px, (
            "OS scroll ticks per step — not screen pixels.\n"
            "Recommended: 5 scroll units at 75% browser PDF zoom gives ~3 captures per page, ideal for AutoStitch.\n"
            "Adjust up or down until consecutive captures overlap cleanly with no gap."
        ))
        tk.Label(row_frame, text="Approx loops:", font=FONT_LABEL).pack(
            side="left", **PAD)
        spin_loops = tk.Spinbox(row_frame, from_=1, to=9999, width=6,
                                textvariable=self._approx_loops,
                                font=FONT_LABEL, state=tk.DISABLED)
        spin_loops.pack(side="left", **PAD)
        tip_loops = tk.Label(row_frame, text="(?)", font=FONT_LABEL, fg="#3498db",
                             cursor="question_arrow")
        tip_loops.pack(side="left", **PAD)
        Tooltip(tip_loops, (
            "Approximate number of scroll-captures needed to cover the whole document.\n"
            "With AutoStitch: ~3 captures per page → 150 loops for a 50-page PDF at 75% zoom with 5 scroll units.\n"
            "At 100% zoom with 4 scroll units: ~4 captures per page → 200 loops for 50 pages.\n"
            "A confirmation dialog appears after this many captures until you press Stop."
        ))

        # collect widgets to enable/disable together
        self._autoscroll_widgets = [rb_v, rb_h, rb_rev, spin_px, spin_loops]

        # ── Section 4: Automate and Relax ─────────────────────────────────────
        self._section(container, "4.  Automate and Relax~")
        row_frame = tk.Frame(container)
        row_frame.pack(fill="x", padx=6)
        self._start_btn = tk.Button(
            row_frame, text="▶  Start Appending", font=FONT_BTN,
            bg="#27ae60", fg="white", width=18,
            command=self._start_appending)
        self._start_btn.pack(side="left", **PAD)
        self._stop_btn = tk.Button(
            row_frame, text="■  Stop Appending", font=FONT_BTN,
            bg="#c0392b", fg="white", width=18,
            command=self._stop_appending, state=tk.DISABLED)
        self._stop_btn.pack(side="left", **PAD)
        self._status_canvas = tk.Canvas(row_frame, width=16, height=16,
                                        highlightthickness=0)
        self._status_canvas.pack(side="left", **PAD)
        self._status_dot = self._status_canvas.create_oval(
            2, 2, 14, 14, fill="#888", outline="")
        self._status_label = tk.Label(row_frame, text="Stopped",
                                      font=FONT_LABEL, fg="#888")
        self._status_label.pack(side="left", **PAD)

        # ── Section: Utilities ────────────────────────────────────────────────
        self._section(container, "Utilities")
        row_frame = tk.Frame(container)
        row_frame.pack(fill="x", padx=6)
        tk.Checkbutton(row_frame, text="Stay Always On Top",
                       variable=self._always_top,
                       font=FONT_LABEL).pack(side="left", **PAD)
        tk.Checkbutton(row_frame, text="Upscale image (2×) before appending",
                       variable=self._upscale,
                       font=FONT_LABEL).pack(side="left", **PAD)

        row_frame = tk.Frame(container)
        row_frame.pack(fill="x", padx=6)
        self._merge_btn = tk.Button(
            row_frame, text="⧉  Merge PDF to Single Page ▾", font=FONT_BTN,
            bg="#8e44ad", fg="white", command=self._merge_to_single_page)
        self._merge_btn.pack(side="left", **PAD)
        tk.Button(row_frame, text="Import Progress", font=FONT_BTN,
                  command=self._import_progress).pack(side="left", **PAD)
        tk.Button(row_frame, text="Export Progress", font=FONT_BTN,
                  command=self._export_progress).pack(side="left", **PAD)

        self._always_top.trace_add("write", lambda *_: None)  # placeholder keep trace count consistent

        self._info_bar = tk.Label(
            container, text="Tip: Win+Shift+S captures to clipboard automatically.",
            font=("Segoe UI", 8), fg="#999", anchor="w")
        self._info_bar.pack(fill="x", padx=6, pady=(4, 0))

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
        save_progress(self)

    def _toggle_autoscroll_ui(self):
        if self._autoscroll_on.get():
            self._autonext_on.set(False)
            self._toggle_autonext_ui_silent()
            for w in self._autoscroll_widgets:
                w.config(state=tk.NORMAL)
        else:
            for w in self._autoscroll_widgets:
                w.config(state=tk.DISABLED)
        save_progress(self)

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

    def _peek_press(self, event):
        if not self._capture_region:
            return
        self.withdraw()
        self.update_idletasks()
        time.sleep(0.15)
        x1, y1, x2, y2 = self._capture_region

        self._peek_top = tk.Toplevel(self)
        self._peek_top.attributes("-fullscreen", True)
        self._peek_top.attributes("-alpha", 0.25)
        self._peek_top.attributes("-topmost", True)
        self._peek_top.configure(bg="black")

        c = tk.Canvas(self._peek_top, bg="black", highlightthickness=0)
        c.pack(fill=tk.BOTH, expand=True)
        c.create_rectangle(x1, y1, x2, y2, outline="#00ff99", width=2, dash=(4, 2))
        tk.Label(c, text="Release to return", bg="black", fg="white",
                 font=("Segoe UI", 13, "bold")).place(relx=0.5, rely=0.04, anchor="center")

    def _peek_release(self, event):
        if hasattr(self, "_peek_top") and self._peek_top:
            self._peek_top.destroy()
            self._peek_top = None
        self.deiconify()

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
            save_progress(self)

    # ── Set click location ─────────────────────────────────────────────────

    def _set_click_location(self):
        if not PYNPUT_OK:
            messagebox.showwarning(
                "pynput missing",
                "Install pynput to capture click location:\n  pip install pynput")
            return

        self._set_status("Click anywhere to set Autonext click position…", color="#888")
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
        self._set_status("Click position captured." if pos else "Cancelled.", color="#888")
        save_progress(self)

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
        self._stitch_frames = []
        self._stop_event.clear()
        self._start_btn.config(state=tk.DISABLED)
        self._stop_btn.config(state=tk.NORMAL)
        self._merge_btn.config(state=tk.DISABLED)
        self._set_status("Running", color="#27ae60", animate=True)

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
                self._set_status("Running (clipboard only – pynput missing)", color="#888")

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
        self._merge_btn.config(state=tk.NORMAL)

        if self._autostitch.get() and len(self._stitch_frames) > 0:
            try:
                stitched = self._smart_stitch_images()
                path = self._pdf_path()
                if os.path.exists(path):
                    os.remove(path)
                append_image_to_pdf(path, stitched)
                self._set_status("AutoStitch complete → 1 page saved", color="#27ae60")
            except Exception as e:
                self._set_status(f"Stitch failed: {e}", color="#e74c3c")
        else:
            self._set_status("Stopped", color="#888")

    # ─────────────────────────────────────────────────────────────────────────
    #  Capture helpers
    # ─────────────────────────────────────────────────────────────────────────

    def _capture_region_image(self) -> Image.Image | None:
        if not self._capture_region:
            self._set_status("No capture area defined.", color="#888")
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

            if self._autostitch.get():
                self._stitch_frames.append(img)
                self._set_status(
                    f"Captured (stitch buffer) → {len(self._stitch_frames)} frames",
                    color="#888"
                )
            else:
                append_image_to_pdf(self._pdf_path(), img)
                self._set_status(
                    f"Page appended → {os.path.basename(self._pdf_path())}",
                    color="#888"
                )

            save_progress(self)
        except Exception as exc:
            self._set_status(f"Error saving: {exc}", color="#e74c3c")

    def _smart_stitch_images(self) -> Image.Image:
        if not self._stitch_frames:
            raise ValueError("No frames available for stitching.")

        direction = self._scroll_direction.get()
        frames = [f.convert("RGB") for f in self._stitch_frames]

        if direction == "horizontal":
            total_w = sum(f.width for f in frames)
            max_h   = max(f.height for f in frames)
            canvas  = Image.new("RGB", (total_w, max_h), (255, 255, 255))
            x = 0
            for f in frames:
                canvas.paste(f, (x, 0))
                x += f.width
            return canvas

        # ── VERTICAL ──────────────────────────────────────────────────────────
        # Strategy: for each consecutive pair (A, B), find exactly how many
        # rows at the BOTTOM of A appear again at the TOP of B, then discard
        # those rows from the top of B before stacking.

        def find_overlap(arr_a, arr_b):
            """
            Returns the number of pixel rows that arr_b shares with the
            bottom of arr_a.  Uses template matching on grayscale strips.
            """
            gray_a = cv2.cvtColor(arr_a, cv2.COLOR_RGB2GRAY)
            gray_b = cv2.cvtColor(arr_b, cv2.COLOR_RGB2GRAY)

            ha, wa = gray_a.shape
            hb, wb = gray_b.shape

            if wa != wb:
                return 0

            # The overlap can be at most half the height of the shorter image
            max_overlap = min(ha, hb) // 2
            if max_overlap < 4:
                return 0

            # Take the bottom strip of A as the search zone,
            # and the top strip of B as the template.
            # We try template heights from large to small and pick the best.
            best_overlap = 0
            best_score   = 0.0

            for frac in (0.5, 0.4, 0.3, 0.2):
                t_h = max(int(min(ha, hb) * frac), 8)
                template = gray_b[:t_h, :]          # top t_h rows of B
                search   = gray_a[ha - t_h * 2:, :] # bottom 2×t_h rows of A

                if search.shape[0] <= template.shape[0]:
                    continue

                res = cv2.matchTemplate(search, template, cv2.TM_CCOEFF_NORMED)
                _, val, _, loc = cv2.minMaxLoc(res)

                if val > best_score:
                    best_score = val
                    # loc[1] = where template was found inside `search`
                    # rows of A below that point = overlap
                    best_overlap = search.shape[0] - loc[1]

            if best_score < 0.70:   # not confident enough → no overlap
                return 0

            # clamp
            return max(0, min(best_overlap, max_overlap))

        arrays = [np.array(f) for f in frames]
        strips = [arrays[0]]

        for i in range(1, len(arrays)):
            ov = find_overlap(strips[-1], arrays[i])
            # discard the overlapping top rows of the incoming frame
            strips.append(arrays[i][ov:] if ov > 0 else arrays[i])

        width     = max(a.shape[1] for a in strips)
        total_h   = sum(a.shape[0] for a in strips)
        canvas    = np.full((total_h, width, 3), 255, dtype=np.uint8)
        y = 0
        for a in strips:
            h, w = a.shape[:2]
            canvas[y:y + h, :w] = a
            y += h

        return Image.fromarray(canvas)

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
        self.after(0, self._set_status, f"Autonext complete – {n} captures saved.", color="#27ae60")

    # ─────────────────────────────────────────────────────────────────────────
    #  Autoscroll loop
    # ─────────────────────────────────────────────────────────────────────────

    def _autoscroll_loop(self):
        import time
        import threading
        from PIL import ImageGrab
        from pynput.mouse import Controller

        approx = self._approx_loops.get()
        if not self._capture_region:
            self.after(0, self._set_status, "No capture region set.", color="#888")
            return

        x1, y1, x2, y2 = self._capture_region
        width  = x2 - x1
        height = y2 - y1

        # Center point for scroll focus
        center_x = (x1 + x2) // 2
        center_y = (y1 + y2) // 2

        # Settings
        direction = self._scroll_direction.get()   # "vertical" / "horizontal"
        reverse   = self._scroll_reverse.get()
        if self._scroll_pixels.get() > 0:
            step_size = self._scroll_pixels.get()
        else:
            step_size = height if direction == "vertical" else width

        delay_capture = 0.25
        delay_scroll  = 0.35

        mouse = Controller()
        iteration = 0

        # Wheel calibration: pixels per wheel unit (approximate, tune if needed)
        PIXELS_PER_TICK = 100.0

        # Floating accumulator for precise stepping
        carry = 0.0

        # Scroll sign
        sign = 1 if reverse else -1

        # Minimize UI
        self.after(0, self.iconify)
        time.sleep(0.4)

        while not self._stop_event.is_set():
            # ---- Capture ----
            x1, y1, x2, y2 = self._capture_region
            bbox = (x1, y1, x2, y2)
            image = ImageGrab.grab(bbox=bbox)
            if image:
                self._save_capture(image)

            iteration += 1

            if self._stop_event.is_set():
                break

            # ---- confirmation after approx loops (UX preserved) ----
            if iteration >= approx:
                confirmed = [None]
                event = threading.Event()

                def _ask():
                    dlg = tk.Toplevel(self)
                    dlg.title("Continue scrolling?")
                    dlg.attributes("-topmost", True)
                    dlg.resizable(False, False)
                    dlg.grab_set()

                    tk.Label(
                        dlg,
                        text=(f"Reached {iteration} loops "
                              f"(approx target was {approx}).\n\n"
                              "Continue capturing?\n\n"
                              "  Enter = Yes, one more then ask again\n"
                              "  Esc   = Stop appending now"),
                        font=("Segoe UI", 10),
                        padx=20, pady=14,
                        justify=tk.LEFT
                    ).pack()

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
                    self.after(0, self.deiconify)
                    return

            time.sleep(delay_capture)

            # ---- Compute scroll ticks (UNCHANGED) ----
            carry += step_size
            ticks_float = carry / PIXELS_PER_TICK
            ticks = int(ticks_float)
            carry -= ticks * PIXELS_PER_TICK

            if ticks == 0:
                ticks = 1

            ticks *= sign

            # ---- Perform scroll (UNCHANGED) ----
            mouse.position = (center_x, center_y)
            time.sleep(0.05)

            if direction == "vertical":
                mouse.scroll(0, ticks)
            else:
                mouse.scroll(ticks, 0)

            time.sleep(delay_scroll)

        # Restore UI
        self.after(0, self.deiconify)
        self.after(0, self._set_status, f"Captured {iteration} frames.", color="#27ae60")

    # ─────────────────────────────────────────────────────────────────────────
    #  Status indicator
    # ─────────────────────────────────────────────────────────────────────────

    def _merge_to_single_page(self):
        try:
            import fitz
        except ImportError:
            messagebox.showerror("Missing library", "PyMuPDF not found.\nRun:  pip install pymupdf")
            return

        use_picker = [False]

        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="Use current PDF", command=lambda: self._do_merge(None))
        menu.add_command(label="Choose a different PDF…", command=lambda: self._do_merge("pick"))
        menu.add_command(label="Merge multiple PDFs into one long page…", command=lambda: self._do_merge("multi"))

        btn = self._merge_btn
        x = btn.winfo_rootx()
        y = btn.winfo_rooty() + btn.winfo_height()
        menu.tk_popup(x, y)

    def _do_merge(self, mode):
        try:
            import fitz
        except ImportError:
            return

        if mode == "multi":
            pdf_paths = filedialog.askopenfilenames(
                title="Select PDFs to merge (in order)", filetypes=[("PDF Files", "*.pdf")])
            if not pdf_paths:
                return
            pdf_paths = list(pdf_paths)
        elif mode == "pick":
            pdf_path = filedialog.askopenfilename(
                title="Select PDF to merge", filetypes=[("PDF Files", "*.pdf")])
            if not pdf_path:
                return
            pdf_paths = [pdf_path]
        else:
            pdf_path = self._pdf_path()
            if not os.path.exists(pdf_path):
                messagebox.showerror("File not found",
                    f"No PDF found at:\n{pdf_path}\n\nCapture some pages first.")
                return
            pdf_paths = [pdf_path]

        self._set_status("Merging pages into single page…", color="#888")
        self.update_idletasks()

        try:
            images = []
            for pdf_path in pdf_paths:
                doc = fitz.open(pdf_path)
                for page in doc:
                    pix = page.get_pixmap(matrix=fitz.Matrix(1, 1))
                    mode_str = "RGBA" if pix.alpha else "RGB"
                    img = Image.frombytes(mode_str, [pix.width, pix.height], pix.samples)
                    images.append(img)
                doc.close()

            # Smart stitch: detect and remove overlapping rows between consecutive images
            rgb_images = [img.convert("RGB") for img in images]
            arrays = [np.array(f) for f in rgb_images]

            def find_overlap_merge(arr_a, arr_b):
                gray_a = cv2.cvtColor(arr_a, cv2.COLOR_RGB2GRAY)
                gray_b = cv2.cvtColor(arr_b, cv2.COLOR_RGB2GRAY)
                ha, wa = gray_a.shape
                hb, wb = gray_b.shape
                if wa != wb:
                    return 0
                max_overlap = min(ha, hb) // 2
                if max_overlap < 4:
                    return 0
                best_overlap = 0
                best_score = 0.0
                for frac in (0.5, 0.4, 0.3, 0.2):
                    t_h = max(int(min(ha, hb) * frac), 8)
                    template = gray_b[:t_h, :]
                    search = gray_a[ha - t_h * 2:, :]
                    if search.shape[0] <= template.shape[0]:
                        continue
                    res = cv2.matchTemplate(search, template, cv2.TM_CCOEFF_NORMED)
                    _, val, _, loc = cv2.minMaxLoc(res)
                    if val > best_score:
                        best_score = val
                        best_overlap = search.shape[0] - loc[1]
                if best_score < 0.70:
                    return 0
                return max(0, min(best_overlap, max_overlap))

            strips = [arrays[0]]
            for i in range(1, len(arrays)):
                ov = find_overlap_merge(arrays[i - 1], arrays[i])
                strips.append(arrays[i][ov:] if ov > 0 else arrays[i])

            width = max(a.shape[1] for a in strips)
            total_h = sum(a.shape[0] for a in strips)
            canvas_arr = np.full((total_h, width, 3), 255, dtype=np.uint8)
            y = 0
            for a in strips:
                h, w = a.shape[:2]
                canvas_arr[y:y + h, :w] = a
                y += h
            combined = Image.fromarray(canvas_arr)

            combined_arr = np.array(combined.convert("RGB"))
            MAX_DIM = 65500
            h, w = combined_arr.shape[:2]
            if h > MAX_DIM or w > MAX_DIM:
                scale = min(MAX_DIM / h, MAX_DIM / w, 1.0)
                new_w = max(1, int(w * scale))
                new_h = max(1, int(h * scale))
                combined_arr = cv2.resize(combined_arr, (new_w, new_h), interpolation=cv2.INTER_AREA)
                combined = Image.fromarray(combined_arr)

            save_path = filedialog.asksaveasfilename(
                title="Save merged single-page PDF",
                defaultextension=".pdf",
                filetypes=[("PDF Files", "*.pdf"), ("PNG Files", "*.png"), ("TIFF Files", "*.tiff")]
            )
            if not save_path:
                self._set_status("Merge cancelled.", color="#888")
                return

            if save_path.lower().endswith(".pdf"):
                combined.convert("RGB").save(save_path, "PDF", resolution=300.0)
            else:
                combined.save(save_path)

            self._set_status(f"Merged → {os.path.basename(save_path)}", color="#27ae60")
        except Exception as exc:
            messagebox.showerror("Merge error", str(exc))
            self._set_status("Merge failed.", color="#e74c3c")

    def _export_progress(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".autocapture",
            filetypes=[("AutoCapture Files", "*.autocapture")]
        )
        if path:
            try:
                save_progress(self, path)
                self._set_status("Exported! → Autosaved", color="#27ae60")
            except Exception:
                self._set_status("Export failed", color="#e74c3c")

    def _import_progress(self):
        path = filedialog.askopenfilename(
            filetypes=[("AutoCapture Files", "*.autocapture")]
        )
        if path:
            try:
                load_progress(self, path)
                self._set_status("Imported! → Autosaved", color="#27ae60")
            except Exception:
                self._set_status("Import failed", color="#e74c3c")

    # ── AUTOSAVE LOGIC ──
    def _schedule_autosave(self, *_):
        """Wait 2s after last change, then save."""
        if self._autosave_job:
            self.after_cancel(self._autosave_job)

        self._autosave_job = self.after(self._AUTOSAVE_DELAY, self._autosave_now)


    def _autosave_now(self):
        self._set_status("Autosaving...", color="#f1c40f", animate=True)

        try:
            save_progress(self)
            self._set_status("Autosaved", color="#27ae60")
        except Exception as e:
            self._set_status(f"Autosave failed", color="#e74c3c")

    def _set_status(self, text: str, color="#888", animate=False):
        self._status_label.config(text=text, fg=color)

        if animate:
            self._animate_status(0)
        else:
            self._status_canvas.itemconfig(self._status_dot, fill=color)

    def _animate_status(self, step):
        if not hasattr(self, "_running") or not self._running:
            return

        colors = ["#f1c40f", "#f39c12", "#f7dc6f"]  # yellow shades
        self._status_canvas.itemconfig(
            self._status_dot,
            fill=colors[step % len(colors)]
        )
        self.after(200, self._animate_status, step + 1)

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