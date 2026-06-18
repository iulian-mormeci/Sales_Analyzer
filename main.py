#!/usr/bin/env python3
"""Sales Analyzer — Redesigned Desktop Dashboard.  Run: python main.py"""

import os, sys, json, math, time, logging, threading
from datetime import datetime, date, timedelta
from pathlib import Path
from typing import Optional, Dict, List, Tuple, Any

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.font as tkfont

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import matplotlib.ticker as mticker

try:
    from tkcalendar import DateEntry
    # tkcalendar uses an overrideredirect Toplevel for its picker, which steals
    # OS-level window focus on macOS and makes the main window unresponsive.
    HAS_CALENDAR = sys.platform != "darwin"
except ImportError:
    HAS_CALENDAR = False

try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# ── Logging ───────────────────────────────────────────────────────────────────
BASE_DIR = Path(os.path.dirname(os.path.abspath(__file__)))
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(BASE_DIR / "app.log", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger(__name__)

try:
    APP_VERSION = (BASE_DIR / "VERSION").read_text().strip()
except Exception:
    APP_VERSION = "1.0.0"

# ── Color system ──────────────────────────────────────────────────────────────
C: Dict[str, str] = {
    "bg":               "#0d0d0d",
    "surface":          "#161616",
    "surface_elevated": "#1f1f1f",
    "accent":           "#6c63ff",
    "accent_warm":      "#ff6b6b",
    "accent_green":     "#43e97b",
    "text_primary":     "#f0f0f0",
    "text_secondary":   "#888888",
    "border":           "#2a2a2a",
    # chart aliases
    "chart_bg":         "#161616",
    "chart_grid":       "#2a2a2a",
}

CHART_COLORS = [
    "#6c63ff","#43e97b","#ff6b6b","#38bdf8","#fbbf24",
    "#a78bfa","#4ecdc4","#fb923c","#f472b6","#34d399",
]

# ── Chart type registry ────────────────────────────────────────────────────────
CHART_TYPES: List[Tuple[str, str, str]] = [
    ("line",    "📈", "Linee"),
    ("bar_v",   "📊", "Barre vert."),
    ("area",    "📉", "Area"),
    ("bar_h",   "▬▬▬", "Barre orizz."),
    ("pie",     "🥧", "Torta"),
    ("scatter", "🔵", "Scatter"),
]

CHART_COMPAT: Dict[str, set] = {
    "trend":      {"line", "area", "bar_v"},
    "top":        {"bar_h", "bar_v", "pie"},
    "comparison": {"bar_v", "line"},
    "geo":        set(),
}

DEFAULT_CHART_TYPES: Dict[str, str] = {
    "trend":      "line",
    "top":        "bar_h",
    "comparison": "bar_v",
}

# ── Column synonyms (unchanged) ────────────────────────────────────────────────
COLUMN_SYNONYMS: Dict[str, List[str]] = {
    "date":    ["data","date","periodo","period","mese","month","data_vendita",
                "data vendita","data_ordine","order_date","giorno","day","timestamp","anno","year"],
    "product": ["prodotto","product","categoria","category","articolo","item",
                "nome_prodotto","nome prodotto","sku","descrizione","description","nome"],
    "revenue": ["ricavo","revenue","fatturato","vendite","importo","amount","totale",
                "total","prezzo","valore","value","incasso","guadagno","entrata","sales",
                "price","ricavi","revenues","importo_totale"],
    "quantity":["quantita","quantity","qty","pezzi","unita","units","quantità",
                "num","numero","count","sold","venduti","q.ta","qta"],
    "channel": ["canale","channel","canale_vendita","sales_channel","tipo_vendita",
                "tipo","type","mezzo","source","origine","modalita","modalità"],
    "geo":     ["citta","city","città","regione","region","provincia","province",
                "luogo","location","area","zona","zone","territorio"],
}

ITALIAN_CITIES: Dict[str, Tuple[float, float]] = {
    "Roma":(41.9028,12.4964),"Milano":(45.4654,9.1859),"Napoli":(40.8518,14.2681),
    "Torino":(45.0703,7.6869),"Palermo":(38.1157,13.3615),"Genova":(44.4056,8.9463),
    "Bologna":(44.4949,11.3426),"Firenze":(43.7696,11.2558),"Bari":(41.1171,16.8719),
    "Catania":(37.5079,15.0830),"Venezia":(45.4408,12.3155),"Verona":(45.4384,10.9916),
    "Messina":(38.1938,15.5540),"Padova":(45.4064,11.8768),"Trieste":(45.6495,13.7768),
    "Brescia":(45.5416,10.2118),"Taranto":(40.4640,17.2470),"Modena":(44.6471,10.9252),
    "Reggio Calabria":(38.1113,15.6474),"Reggio Emilia":(44.6989,10.6297),
    "Perugia":(43.1121,12.3888),"Ravenna":(44.4184,12.2035),"Livorno":(43.5485,10.3106),
    "Cagliari":(39.2238,9.1217),"Foggia":(41.4600,15.5440),"Rimini":(44.0678,12.5695),
    "Salerno":(40.6824,14.7681),"Ferrara":(44.8381,11.6198),"Sassari":(40.7259,8.5556),
    "Bergamo":(45.6983,9.6773),
}

# ── Runtime fonts (populated after Tk() init) ─────────────────────────────────
UI_FONT   = "Segoe UI"
MONO_FONT = "Courier New"

def _detect_fonts() -> None:
    global UI_FONT, MONO_FONT
    fam = set(tkfont.families())
    UI_FONT = next(
        (f for f in ["Inter","SF Pro Display",".AppleSystemUIFont",
                     "Segoe UI","Helvetica Neue","Arial"] if f in fam),
        "TkDefaultFont"
    )
    MONO_FONT = next(
        (f for f in ["JetBrains Mono","Fira Code","Cascadia Code",
                     "Consolas","Courier New","Monaco"] if f in fam),
        "TkFixedFont"
    )

# ── Math / color helpers ───────────────────────────────────────────────────────
def _ease_quint(t: float) -> float:
    """Quintic ease-out — snappy start, soft stop."""
    return 1 - (1 - t) ** 5

def _ease_quad(t: float) -> float:
    return 1 - (1 - t) ** 2

def _ease_cubic(t: float) -> float:
    return 1 - (1 - t) ** 3

def _hex_to_rgb(h: str) -> Tuple[int, int, int]:
    h = h.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)

def _lerp_color(c1: str, c2: str, t: float) -> str:
    r1, g1, b1 = _hex_to_rgb(c1)
    r2, g2, b2 = _hex_to_rgb(c2)
    return "#{:02x}{:02x}{:02x}".format(
        int(r1 + (r2 - r1) * t),
        int(g1 + (g2 - g1) * t),
        int(b1 + (b2 - b1) * t),
    )

def _bar_gradient(n: int) -> List[str]:
    """n colors interpolated from accent to accent_warm."""
    return [_lerp_color(C["accent"], C["accent_warm"], i / max(n - 1, 1))
            for i in range(n)]

def _fmt_currency(v: float) -> str:
    return f"€{v:,.0f}"

def _fmt_count(v: float) -> str:
    return f"{int(v):,}"


# ═══════════════════════════════════════════════════════════════════════════════
# Core logic (unchanged)
# ═══════════════════════════════════════════════════════════════════════════════

class ConfigManager:
    def __init__(self):
        self._path = BASE_DIR / "config.json"
        self._data: Dict[str, Any] = {}
        try:
            if self._path.exists():
                self._data = json.loads(self._path.read_text())
        except Exception as e:
            log.warning(f"Config load: {e}")

    def get(self, key, default=None):
        return self._data.get(key, default)

    def set(self, key, value):
        self._data[key] = value
        tmp = self._path.with_suffix(".json.tmp")
        try:
            tmp.write_text(json.dumps(self._data, indent=2))
            os.replace(tmp, self._path)
        except Exception as e:
            log.warning(f"Config save: {e}")
            try:
                tmp.unlink(missing_ok=True)
            except Exception:
                pass


class ColumnMapper:
    def auto_map(self, columns: List[str]) -> Dict[str, Optional[str]]:
        cols_lower = {c: c.lower().strip().replace(" ", "_") for c in columns}
        result: Dict[str, Optional[str]] = {role: None for role in COLUMN_SYNONYMS}
        for role, synonyms in COLUMN_SYNONYMS.items():
            for col, col_l in cols_lower.items():
                for syn in synonyms:
                    if col_l == syn.replace(" ", "_") or col_l.startswith(syn):
                        result[role] = col
                        break
                if result[role]:
                    break
        return result

    def is_complete(self, mapping: Dict[str, Optional[str]]) -> bool:
        return all(mapping.get(r) for r in ("date", "product", "revenue"))


# ═══════════════════════════════════════════════════════════════════════════════
# Animation helpers
# ═══════════════════════════════════════════════════════════════════════════════

def _widget_raise(w: tk.Widget) -> None:
    """Raise w in the window stacking order.

    tk.Canvas overrides both lift() and tkraise() with tag_raise(), so neither
    works when called on a Canvas instance.  The raw Tcl 'raise' command always
    operates on the widget path, not on canvas item tags.
    """
    w.tk.call("raise", w._w)


def _run_anim(widget: tk.Widget, duration_ms: int, ease_fn,
              on_step, on_done=None) -> None:
    """
    Generic after()-based animation driver.
    Calls on_step(t_eased) every ~16 ms, then on_done().
    """
    t0 = time.monotonic()
    dur = duration_ms / 1000

    def _tick():
        raw = min((time.monotonic() - t0) / dur, 1.0)
        on_step(ease_fn(raw))
        if raw < 1.0:
            widget.after(16, _tick)
        elif on_done:
            on_done()

    _tick()


def _fade_overlay(content: tk.Frame, callback, duration_ms: int = 120) -> None:
    """
    Fade content to black, call callback, fade back in.
    Falls back to instant swap when PIL is unavailable or widget too small.
    """
    content.update_idletasks()
    w, h = content.winfo_width(), content.winfo_height()
    if not HAS_PIL or w < 2 or h < 2:
        callback()
        return

    ov = tk.Canvas(content, bg=C["bg"], highlightthickness=0)
    ov.place(x=0, y=0, width=w, height=h)
    _widget_raise(ov)
    ref: list = [None]
    iid: list = [None]
    rr, rg, rb = _hex_to_rgb(C["bg"])

    def _draw(alpha: int) -> None:
        img = Image.new("RGBA", (w, h), (rr, rg, rb, alpha))
        photo = ImageTk.PhotoImage(img)
        ref[0] = photo
        if iid[0] is None:
            iid[0] = ov.create_image(0, 0, image=photo, anchor="nw")
        else:
            ov.itemconfig(iid[0], image=photo)

    def _phase2() -> None:
        t0 = time.monotonic()
        def _step(te):
            _draw(int(255 * (1 - te)))
        def _done():
            ov.destroy()
        _run_anim(ov, duration_ms, _ease_cubic, _step, _done)

    def _phase1_done() -> None:
        callback()
        ov.after(16, _phase2)

    _run_anim(ov, duration_ms, _ease_cubic,
              lambda te: _draw(int(255 * te)),
              _phase1_done)


def _wipe_reveal(parent: tk.Frame, duration_ms: int = 350) -> None:
    """
    Reveals a newly-drawn chart with a left-to-right wipe.
    Places a solid cover that retreats rightward.
    """
    parent.update_idletasks()
    w = parent.winfo_width()
    h = parent.winfo_height()
    if w < 2:
        return

    cover = tk.Canvas(parent, bg=C["bg"], highlightthickness=0)
    cover.place(x=0, y=0, width=w, height=h)
    _widget_raise(cover)

    def _step(te: float) -> None:
        x = int(w * te)
        cover.place(x=x, y=0, width=max(1, w - x), height=h)

    _run_anim(cover, duration_ms, _ease_cubic, _step,
              on_done=cover.destroy)


# ═══════════════════════════════════════════════════════════════════════════════
# Reusable widgets
# ═══════════════════════════════════════════════════════════════════════════════

def _hover(w: tk.Widget, normal: str, hot: str) -> None:
    """Attach subtle background hover to any widget."""
    w.bind("<Enter>", lambda _: w.config(bg=hot))
    w.bind("<Leave>", lambda _: w.config(bg=normal))


class _Btn(tk.Label):
    """Flat label-button with hover, cursor, and click binding."""

    def __init__(self, parent, text, command, bg=None, fg=None,
                 font_args=None, padx=12, pady=5, cursor="hand2", **kw):
        _bg = bg or C["surface"]
        _fg = fg or C["text_primary"]
        _font = font_args or (UI_FONT, 10)
        super().__init__(parent, text=text, bg=_bg, fg=_fg,
                         font=_font, padx=padx, pady=pady,
                         cursor=cursor, **kw)
        self._bg = _bg
        self._hot = _lerp_color(_bg, "#ffffff", 0.08)
        self.bind("<Button-1>", lambda _: command())
        _hover(self, _bg, self._hot)

    def set_active(self, active: bool) -> None:
        if active:
            self.config(bg=C["accent"], fg="#ffffff")
            self._bg = C["accent"]
            self._hot = _lerp_color(C["accent"], "#ffffff", 0.1)
        else:
            self.config(bg=C["surface"], fg=C["text_secondary"])
            self._bg = C["surface"]
            self._hot = _lerp_color(C["surface"], "#ffffff", 0.06)


class PillGroup(tk.Frame):
    """Horizontal pill toggle — selected pill gets accent bg."""

    def __init__(self, parent, options: List[str],
                 initial: str = "", callback=None, bg=None, **kw):
        _bg = bg or C["bg"]
        super().__init__(parent, bg=_bg, **kw)
        self._sel = initial or options[0]
        self._cb  = callback
        self._pills: Dict[str, tk.Label] = {}

        wrap  = tk.Frame(self, bg=C["border"])
        wrap.pack()
        inner = tk.Frame(wrap, bg=C["surface"])
        inner.pack(padx=1, pady=1)

        for opt in options:
            lbl = tk.Label(inner, text=opt, padx=11, pady=4,
                           cursor="hand2", font=(UI_FONT, 10))
            lbl.pack(side="left")
            lbl.bind("<Button-1>", lambda _, o=opt: self._pick(o))
            self._pills[opt] = lbl

        self._redraw()

    def _pick(self, opt: str) -> None:
        self._sel = opt
        self._redraw()
        if self._cb:
            self._cb(opt)

    def _redraw(self) -> None:
        for opt, lbl in self._pills.items():
            if opt == self._sel:
                lbl.config(bg=C["accent"], fg="#ffffff")
            else:
                lbl.config(bg=C["surface"], fg=C["text_secondary"])
                _hover(lbl, C["surface"], C["surface_elevated"])

    def get(self) -> str:
        return self._sel

    def set(self, v: str) -> None:
        if v in self._pills:
            self._sel = v
            self._redraw()


class ChannelDropdown(tk.Frame):
    """Dropdown button with an in-window checklist panel.

    The panel is a plain tk.Frame placed inside the root window (not a
    Toplevel) so it never steals OS-level window focus — which on macOS
    causes the main window to become unresponsive to clicks until
    minimized/restored.
    """

    def __init__(self, parent, bg=None, **kw):
        _bg = bg or C["bg"]
        super().__init__(parent, bg=_bg, **kw)
        self.options: List[str] = []
        self.vars: Dict[str, tk.BooleanVar] = {}
        self._panel: Optional[tk.Frame] = None
        self._unbind_id: Optional[str] = None

        self._btn = tk.Label(self, text="Canale  ▾", bg=C["surface"],
                             fg=C["text_secondary"], font=(UI_FONT, 10),
                             padx=12, pady=5, cursor="hand2")
        self._btn.pack()
        self._btn.bind("<Button-1>", lambda _: self._toggle())
        _hover(self._btn, C["surface"], C["surface_elevated"])

    def set_options(self, options: List[str]) -> None:
        self.options = sorted(options)
        self.vars = {o: tk.BooleanVar(value=True) for o in self.options}
        self._update_label()

    def get_selected(self) -> List[str]:
        return [k for k, v in self.vars.items() if v.get()]

    def _update_label(self) -> None:
        sel = self.get_selected()
        n = len(self.options)
        if n == 0 or len(sel) == n:
            self._btn.config(text="Canale  ▾")
        elif len(sel) == 0:
            self._btn.config(text="Canale (nessuno)  ▾")
        else:
            self._btn.config(text=f"Canale ({len(sel)}/{n})  ▾")

    def _toggle(self) -> None:
        if self._panel and self._panel.winfo_exists():
            self._close()
        else:
            self._open()

    def _open(self) -> None:
        self._btn.update_idletasks()
        root = self.winfo_toplevel()

        # Compute position relative to root window origin
        bx = self._btn.winfo_rootx() - root.winfo_rootx()
        by = (self._btn.winfo_rooty() - root.winfo_rooty()
              + self._btn.winfo_height() + 4)

        # Build panel as a child of root (no Toplevel = no focus theft)
        panel = tk.Frame(root, bg=C["border"])
        panel.place(x=bx, y=by)
        self._panel = panel

        inner = tk.Frame(panel, bg=C["surface_elevated"])
        inner.pack(padx=1, pady=1)

        def _all():
            all_on = all(v.get() for v in self.vars.values())
            for v in self.vars.values():
                v.set(not all_on)
            self._update_label()

        sel_lbl = tk.Label(inner, text="Seleziona tutti",
                           bg=C["surface_elevated"], fg=C["accent"],
                           font=(UI_FONT, 9), padx=10, pady=5, cursor="hand2")
        sel_lbl.pack(fill="x")
        sel_lbl.bind("<Button-1>", lambda _: _all())

        tk.Frame(inner, bg=C["border"], height=1).pack(fill="x")

        for opt in self.options:
            row = tk.Frame(inner, bg=C["surface_elevated"])
            row.pack(fill="x", padx=6, pady=2)
            tk.Checkbutton(row, text=opt, variable=self.vars[opt],
                           bg=C["surface_elevated"], fg=C["text_primary"],
                           selectcolor=C["accent"],
                           activebackground=C["surface_elevated"],
                           font=(UI_FONT, 10), command=self._update_label,
                           padx=4).pack(anchor="w")

        # Raise the panel above all sibling widgets
        panel.lift()

        # Close on any click outside the panel
        def _outside(ev):
            if not (self._panel and self._panel.winfo_exists()):
                self._unbind_outside()
                return
            try:
                px, py = self._panel.winfo_rootx(), self._panel.winfo_rooty()
                pw, ph = self._panel.winfo_width(), self._panel.winfo_height()
                inside = (px <= ev.x_root <= px + pw and
                          py <= ev.y_root <= py + ph)
            except Exception:
                inside = False
            if not inside:
                self._close()

        self._unbind_id = root.bind("<Button-1>", _outside, add="+")

    def _close(self) -> None:
        if self._panel and self._panel.winfo_exists():
            self._panel.destroy()
        self._panel = None
        self._unbind_outside()

    def _unbind_outside(self) -> None:
        if self._unbind_id is not None:
            try:
                self.winfo_toplevel().unbind("<Button-1>", self._unbind_id)
            except Exception:
                pass
            self._unbind_id = None


class LoadingPlaceholder(tk.Frame):
    """Pulsing skeleton screen while Excel is loading.

    Inherits tk.Frame (not Canvas) so that .lift()/.tkraise() use the widget
    stacking method, not Canvas.tag_raise which requires a tagOrId argument.
    The drawing canvas is a child widget.
    """

    _RECTS = [
        (0.04, 0.06, 0.92, 0.07),
        (0.04, 0.18, 0.92, 0.55),
        (0.04, 0.80, 0.27, 0.06),
        (0.35, 0.80, 0.27, 0.06),
        (0.66, 0.80, 0.27, 0.06),
    ]

    def __init__(self, parent, **kw):
        super().__init__(parent, bg=C["bg"], **kw)
        self._canvas = tk.Canvas(self, bg=C["bg"], highlightthickness=0)
        self._canvas.pack(fill="both", expand=True)
        self._rects: List[int] = []
        self._phase = 0.0
        self._job: Optional[str] = None
        self._canvas.bind("<Configure>", self._rebuild)
        self._animate()

    def _rebuild(self, _=None) -> None:
        self._canvas.delete("all")
        self._rects.clear()
        w, h = self._canvas.winfo_width(), self._canvas.winfo_height()
        if w < 2 or h < 2:
            return
        for xr, yr, wr, hr in self._RECTS:
            rid = self._canvas.create_rectangle(
                int(xr*w), int(yr*h), int((xr+wr)*w), int((yr+hr)*h),
                fill=C["surface"], outline="",
            )
            self._rects.append(rid)

    def _animate(self) -> None:
        self._phase = (self._phase + 0.07) % (2 * math.pi)
        t = (1 + math.sin(self._phase)) / 2
        col = _lerp_color(C["surface"], C["surface_elevated"], t)
        for rid in self._rects:
            try:
                self._canvas.itemconfig(rid, fill=col)
            except Exception:
                pass
        self._job = self.after(40, self._animate)   # ~25 fps

    def destroy(self) -> None:
        if self._job:
            try:
                self.after_cancel(self._job)
            except Exception:
                pass
        super().destroy()


class DropZone(tk.Frame):
    """Centered empty state shown before any file is loaded."""

    def __init__(self, parent, on_browse, **kw):
        super().__init__(parent, bg=C["bg"], **kw)
        self._build(on_browse)

    def _build(self, on_browse) -> None:
        inner = tk.Frame(self, bg=C["bg"])
        inner.place(relx=0.5, rely=0.5, anchor="center")

        # Dashed border box (simulated with Canvas)
        box = tk.Canvas(inner, width=320, height=220,
                        bg=C["bg"], highlightthickness=0)
        box.pack()
        # Draw dashed border rectangle
        box.create_rectangle(2, 2, 318, 218, outline=C["border"],
                              width=1, dash=(6, 4))

        tk.Label(inner, text="📊", bg=C["bg"], fg=C["text_secondary"],
                 font=(UI_FONT, 36)).pack(pady=(8, 4))
        tk.Label(inner, text="Carica un file Excel per iniziare",
                 bg=C["bg"], fg=C["text_primary"],
                 font=(UI_FONT, 13, "bold")).pack()
        tk.Label(inner,
                 text="Formati supportati: .xlsx  .xls  .xlsm",
                 bg=C["bg"], fg=C["text_secondary"],
                 font=(UI_FONT, 10)).pack(pady=(2, 14))

        browse = tk.Label(inner, text="  Sfoglia file  ",
                          bg=C["accent"], fg="#ffffff",
                          font=(UI_FONT, 11, "bold"),
                          padx=20, pady=8, cursor="hand2")
        browse.pack()
        browse.bind("<Button-1>", lambda _: on_browse())
        _hover(browse, C["accent"],
               _lerp_color(C["accent"], "#ffffff", 0.12))


# ═══════════════════════════════════════════════════════════════════════════════
# Toast notification
# ═══════════════════════════════════════════════════════════════════════════════

class Toast(tk.Frame):
    """Transient bottom-right notification — auto-dismisses after duration_ms."""

    def __init__(self, root: tk.Tk, message: str, duration_ms: int = 2000):
        super().__init__(root, bg=C["surface_elevated"],
                         highlightthickness=1,
                         highlightbackground=C["border"])
        tk.Label(self, text=message, bg=C["surface_elevated"],
                 fg=C["text_primary"], font=(UI_FONT, 10),
                 padx=16, pady=8).pack()

        root.update_idletasks()
        rw = root.winfo_width()
        rh = root.winfo_height()
        self.update_idletasks()
        w = self.winfo_reqwidth()
        h = self.winfo_reqheight()
        self.place(x=rw - w - 20, y=rh - h - 20)
        self.lift()
        self.after(duration_ms, self._fade_out)

    def _fade_out(self) -> None:
        if not HAS_PIL:
            self.destroy()
            return
        self.destroy()


# ═══════════════════════════════════════════════════════════════════════════════
# Column Config Dialog  (user-driven mapping, Feature 1)
# ═══════════════════════════════════════════════════════════════════════════════

class ColumnConfigDialog(tk.Toplevel):
    """
    Full-featured column mapping editor — 420×500 dark modal.
    result = new mapping dict when Applica pressed, None otherwise.
    reset_to_auto = True when user clicked Reset automatico.
    """

    _ROLES = [
        ("date",     "Data / Periodo",    True),
        ("product",  "Prodotto",          True),
        ("revenue",  "Fatturato",         True),
        ("quantity", "Quantità",          False),
        ("channel",  "Canale",            False),
        ("geo",      "Città / Regione",   False),
    ]

    def __init__(self, parent, columns: List[str],
                 mapping: Dict[str, Optional[str]],
                 mapper: "ColumnMapper"):
        super().__init__(parent)
        self.title("Configura Colonne")
        self.configure(bg=C["bg"])
        self.resizable(False, False)
        self.grab_set()
        self.geometry("420x520")
        self.result: Optional[Dict] = None
        self.reset_to_auto = False
        self._columns = columns
        self._mapper  = mapper
        self._combos: Dict[str, ttk.Combobox] = {}
        self._err_lbls: Dict[str, tk.Label]   = {}

        self._center(parent)
        self._build(mapping)

    def _center(self, parent) -> None:
        self.update_idletasks()
        pw = parent.winfo_width(); ph = parent.winfo_height()
        px = parent.winfo_rootx(); py = parent.winfo_rooty()
        w = 420; h = 520
        self.geometry(f"{w}x{h}+{px+(pw-w)//2}+{py+(ph-h)//2}")

    def _build(self, mapping: Dict) -> None:
        # Title
        tk.Label(self, text="Configura Colonne", bg=C["bg"],
                 fg=C["text_primary"], font=(UI_FONT, 13, "bold")).pack(
            pady=(20, 4), padx=24, anchor="w")
        tk.Label(self,
                 text="Associa le colonne del tuo Excel ai campi dell'app.",
                 bg=C["bg"], fg=C["text_secondary"],
                 font=(UI_FONT, 10)).pack(padx=24, anchor="w")
        tk.Frame(self, bg=C["border"], height=1).pack(fill="x", padx=0, pady=12)

        # Style comboboxes
        sty = ttk.Style()
        sty.theme_use("default")
        sty.configure("CC.TCombobox",
                       fieldbackground=C["surface_elevated"],
                       background=C["surface_elevated"],
                       foreground=C["text_primary"],
                       selectbackground=C["accent"],
                       arrowcolor=C["text_secondary"],
                       bordercolor=C["border"])
        sty.map("CC.TCombobox",
                fieldbackground=[("readonly", C["surface_elevated"])],
                foreground=[("readonly", C["text_primary"])])

        choices = ["— non usare —"] + list(self._columns)

        # Scrollable form
        form = tk.Frame(self, bg=C["bg"])
        form.pack(fill="x", padx=24)

        for role, label, required in self._ROLES:
            hint = "" if required else "  (opzionale)"
            row = tk.Frame(form, bg=C["bg"])
            row.pack(fill="x", pady=(6, 0))

            tk.Label(row, text=label + hint, bg=C["bg"],
                     fg=C["text_primary"] if required else C["text_secondary"],
                     font=(UI_FONT, 10), anchor="w",
                     width=20).pack(side="left")

            cb = ttk.Combobox(row, values=choices, state="readonly",
                               width=22, style="CC.TCombobox")
            cur = mapping.get(role)
            cb.set(cur if cur and cur in self._columns else "— non usare —")
            cb.pack(side="left")
            self._combos[role] = cb

            err = tk.Label(form, text="", bg=C["bg"],
                           fg=C["accent_warm"], font=(UI_FONT, 9))
            err.pack(anchor="e", pady=0)
            self._err_lbls[role] = err

        tk.Frame(self, bg=C["border"], height=1).pack(fill="x", pady=12)

        # Buttons
        btns = tk.Frame(self, bg=C["bg"])
        btns.pack(pady=(0, 20))

        _Btn(btns, "Applica", self._apply,
             bg=C["accent"], fg="#ffffff", padx=16, pady=7).pack(
            side="left", padx=6)
        _Btn(btns, "Annulla", self.destroy,
             bg=C["surface"], padx=16, pady=7).pack(side="left", padx=6)
        _Btn(btns, "Reset automatico", self._reset_auto,
             bg=C["surface_elevated"], fg=C["text_secondary"],
             padx=12, pady=7).pack(side="left", padx=6)

    def _apply(self) -> None:
        # Clear previous errors
        for lbl in self._err_lbls.values():
            lbl.config(text="")

        result = {r: (None if cb.get().startswith("—") else cb.get())
                  for r, cb in self._combos.items()}

        errors = []
        if not result.get("date"):
            self._err_lbls["date"].config(
                text="⚠  La colonna Data è obbligatoria")
            errors.append("date")
        if not result.get("revenue") and not result.get("quantity"):
            self._err_lbls["revenue"].config(
                text="⚠  Almeno Fatturato o Quantità è richiesto")
            errors.append("revenue")

        if errors:
            return

        self.result = result
        self.destroy()

    def _reset_auto(self) -> None:
        auto = self._mapper.auto_map(self._columns)
        for role, cb in self._combos.items():
            cur = auto.get(role)
            cb.set(cur if cur and cur in self._columns else "— non usare —")
        for lbl in self._err_lbls.values():
            lbl.config(text="")
        self.reset_to_auto = True


# ═══════════════════════════════════════════════════════════════════════════════
# Chart Type Panel  (Feature 2)
# ═══════════════════════════════════════════════════════════════════════════════

class ChartTypePanel(tk.Frame):
    """
    Floating in-window panel for selecting chart type.
    Placed as a child of the root window (not Toplevel) to avoid macOS focus issues.
    """

    # View labels used in the panel subtitle
    _VIEW_LABELS = {
        "trend":      "Trend",
        "top":        "Top Prodotti",
        "comparison": "Confronto",
        "geo":        "Mappa",
    }

    def __init__(self, root: tk.Tk, anchor_widget: tk.Widget,
                 view_key: str, current_type: str,
                 on_select, on_close):
        super().__init__(root, bg=C["border"])
        self._on_select  = on_select
        self._on_close   = on_close
        self._view_key   = view_key
        self._unbind_id: Optional[str] = None
        self._picked     = False      # guard against double-fire on propagation

        body = tk.Frame(self, bg=C["surface_elevated"])
        body.pack(padx=1, pady=1)

        compat       = CHART_COMPAT.get(view_key, set())
        geo_disabled = (view_key == "geo")
        view_lbl     = self._VIEW_LABELS.get(view_key, view_key)

        # ── Header ────────────────────────────────────────────────────────────
        hdr = tk.Frame(body, bg=C["surface_elevated"])
        hdr.pack(fill="x", padx=12, pady=(10, 4))
        tk.Label(hdr, text="Tipo di grafico", bg=C["surface_elevated"],
                 fg=C["text_primary"], font=(UI_FONT, 10, "bold")).pack(side="left")

        # Subtitle: show which view is active and what's available
        if geo_disabled:
            sub = "La mappa non supporta questo setting"
            sub_col = C["accent_warm"]
        else:
            avail_names = [l for k, _, l in CHART_TYPES if k in compat]
            sub = f"Vista: {view_lbl}  ·  disponibili: {', '.join(avail_names)}"
            sub_col = C["text_secondary"]

        tk.Label(body, text=sub, bg=C["surface_elevated"],
                 fg=sub_col, font=(UI_FONT, 8),
                 wraplength=260, justify="left").pack(
            anchor="w", padx=12, pady=(0, 8))

        tk.Frame(body, bg=C["border"], height=1).pack(fill="x", padx=0)

        # ── 2-column card grid ────────────────────────────────────────────────
        grid = tk.Frame(body, bg=C["surface_elevated"])
        grid.pack(padx=10, pady=10)

        for i, (ct, icon, label) in enumerate(CHART_TYPES):
            available = (ct in compat) and not geo_disabled
            selected  = (ct == current_type) and available
            card = self._make_card(grid, ct, icon, label, available, selected)
            card.grid(row=i // 2, column=i % 2, padx=5, pady=5, sticky="nsew")

        # ── Position clamped inside root ──────────────────────────────────────
        self.place(x=0, y=0)        # place first so tkinter sizes it
        root.update_idletasks()

        rw = root.winfo_width();  rh = root.winfo_height()
        pw = self.winfo_reqwidth(); ph = self.winfo_reqheight()

        ax = anchor_widget.winfo_rootx() - root.winfo_rootx()
        ay = (anchor_widget.winfo_rooty() - root.winfo_rooty()
              + anchor_widget.winfo_height() + 4)

        if ay + ph > rh:            # flip above anchor if it would overflow bottom
            ay = (anchor_widget.winfo_rooty() - root.winfo_rooty()) - ph - 4
        ax = max(4, min(ax, rw - pw - 4))   # clamp left/right
        ay = max(4, ay)                      # clamp top

        self.place(x=ax, y=ay)
        self.lift()

        # ── Close on outside click ────────────────────────────────────────────
        def _outside(ev):
            if not self.winfo_exists():
                self._unbind()
                return
            try:
                px = self.winfo_rootx(); py = self.winfo_rooty()
                pw = self.winfo_width(); ph = self.winfo_height()
                inside = (px <= ev.x_root <= px + pw and
                          py <= ev.y_root <= py + ph)
            except Exception:
                inside = False
            if not inside:
                self._close()

        self._unbind_id = root.bind("<Button-1>", _outside, add="+")

    def _make_card(self, parent, ct, icon, label,
                   available: bool, selected: bool) -> tk.Frame:
        """
        Single chart-type card.  Uses highlightthickness for the border so
        tkinter's geometry manager sees the correct widget size without needing
        a separate outer/inner frame pair.
        """
        if selected:
            bg   = _lerp_color(C["accent"], C["surface_elevated"], 0.82)
            hl   = C["accent"]
            hl_w = 2
        else:
            bg   = C["surface"]
            hl   = C["border"]
            hl_w = 1

        card = tk.Frame(parent, bg=bg, width=114, height=68,
                        highlightthickness=hl_w,
                        highlightbackground=hl,
                        highlightcolor=hl)
        card.pack_propagate(False)   # keep fixed size regardless of children

        dim     = not available
        fg_icon = _lerp_color(C["text_secondary"], C["bg"], 0.5) if dim else C["text_primary"]
        fg_lbl  = _lerp_color(C["text_secondary"], C["bg"], 0.5) if dim else C["text_secondary"]

        icon_lbl = tk.Label(card, text=icon, bg=bg, fg=fg_icon,
                            font=(UI_FONT, 17))
        icon_lbl.place(relx=0.5, rely=0.34, anchor="center")

        text_lbl = tk.Label(card, text=label, bg=bg, fg=fg_lbl,
                            font=(UI_FONT, 8))
        text_lbl.place(relx=0.5, rely=0.74, anchor="center")

        badge_text = "✓" if selected else ("✗" if dim else "")
        badge_col  = (C["accent"] if selected
                      else _lerp_color(C["accent_warm"], C["bg"], 0.35))
        if badge_text:
            tk.Label(card, text=badge_text, bg=bg, fg=badge_col,
                     font=(UI_FONT, 7, "bold")).place(
                relx=0.88, rely=0.14, anchor="center")

        if available:
            hot = _lerp_color(bg, "#ffffff", 0.07)
            widgets = [card, icon_lbl, text_lbl]
            for w in widgets:
                w.config(cursor="hand2")
                # bind on every widget so click registers wherever the user hits
                w.bind("<Button-1>", lambda _e, c=ct: self._pick(c))
                if not selected:
                    _hover(w, bg, hot)

        return card

    def _pick(self, ct: str) -> None:
        # Guard: tkinter may propagate one click to multiple bound widgets;
        # only handle the first one.
        if self._picked:
            return
        self._picked = True
        self._on_select(self._view_key, ct)
        self._close()

    def _close(self) -> None:
        self._unbind()
        self._on_close()
        if self.winfo_exists():
            self.destroy()

    def _unbind(self) -> None:
        if self._unbind_id is not None:
            try:
                self.winfo_toplevel().unbind("<Button-1>", self._unbind_id)
            except Exception:
                pass
            self._unbind_id = None


# ═══════════════════════════════════════════════════════════════════════════════
# KPI strip
# ═══════════════════════════════════════════════════════════════════════════════

class KPICard(tk.Frame):
    """Single KPI card with count-up animation."""

    def __init__(self, parent, icon: str, label: str, **kw):
        super().__init__(parent, bg=C["surface"], **kw)
        # 3-px left accent border
        tk.Frame(self, bg=C["accent"], width=3).pack(side="left", fill="y")

        body = tk.Frame(self, bg=C["surface"], padx=14, pady=10)
        body.pack(side="left", fill="both", expand=True)

        # Icon + label row
        top = tk.Frame(body, bg=C["surface"])
        top.pack(anchor="w")
        tk.Label(top, text=icon, bg=C["surface"],
                 fg=C["text_secondary"], font=(UI_FONT, 13)).pack(side="left")
        tk.Label(top, text=f"  {label}", bg=C["surface"],
                 fg=C["text_secondary"], font=(UI_FONT, 11)).pack(side="left")

        self._val = tk.Label(body, text="—", bg=C["surface"],
                             fg=C["text_primary"],
                             font=(MONO_FONT, 16, "bold"))
        self._val.pack(anchor="w", pady=(3, 1))

        self._delta = tk.Label(body, text="", bg=C["surface"],
                               fg=C["text_secondary"], font=(UI_FONT, 11))
        self._delta.pack(anchor="w")

    def animate(self, value, fmt_fn, delta_str: str = "",
                delta_color: str = "") -> None:
        """Count value up from 0 to target over 800 ms (ease-out quad)."""
        _dc = delta_color or C["text_secondary"]
        if isinstance(value, str) or not isinstance(value, (int, float)):
            self._val.config(text=str(value))
            self._delta.config(text=delta_str, fg=_dc)
            return
        if value == 0:
            self._val.config(text=fmt_fn(0))
            self._delta.config(text=delta_str, fg=_dc)
            return

        target = float(value)

        def _step(te: float) -> None:
            self._val.config(text=fmt_fn(target * te))

        def _done() -> None:
            self._val.config(text=fmt_fn(target))
            self._delta.config(text=delta_str, fg=_dc)

        _run_anim(self, 800, _ease_quad, _step, _done)


class KPIStrip(tk.Frame):
    """Row of 4 KPI cards."""

    _DEFS = [
        ("₿",  "Ricavo Totale"),
        ("📦", "Unità Vendute"),
        ("⌀",  "Ordine Medio"),
        ("📡", "Canale Top"),
    ]

    def __init__(self, parent, **kw):
        super().__init__(parent, bg=C["bg"], **kw)
        self._cards: List[KPICard] = []
        for i, (icon, label) in enumerate(self._DEFS):
            card = KPICard(self, icon, label)
            card.pack(side="left", fill="both", expand=True,
                      padx=(0 if i == 0 else 8, 0), pady=0)
            self._cards.append(card)

    def update(self, kpis) -> None:
        """kpis: list of (icon,label,value,delta_pct,fmt_type) tuples."""
        for card, (_icon, _label, value, delta_pct, fmt_type) in \
                zip(self._cards, kpis):
            if fmt_type == "currency":
                fmt_fn = _fmt_currency
            elif fmt_type == "count":
                fmt_fn = _fmt_count
            else:
                fmt_fn = str

            if delta_pct is not None and isinstance(delta_pct, float):
                sign  = "▲" if delta_pct >= 0 else "▼"
                dcol  = C["accent_green"] if delta_pct >= 0 else C["accent_warm"]
                dstr  = f"{sign} {abs(delta_pct):.1f}%"
            else:
                dcol = C["text_secondary"]
                dstr = ""

            if fmt_type == "channel":
                # value is channel name; delta_pct is % share
                card.animate(str(value), str,
                             f"{delta_pct:.0f}% quota" if delta_pct else "",
                             C["text_secondary"])
            else:
                card.animate(value, fmt_fn, dstr, dcol)


# ═══════════════════════════════════════════════════════════════════════════════
# TopBar
# ═══════════════════════════════════════════════════════════════════════════════

class TopBar(tk.Frame):
    """48-px top bar: logo | filename  ··  version badge | load button."""

    def __init__(self, parent, on_load, **kw):
        super().__init__(parent, bg=C["bg"], height=48, **kw)
        self.pack_propagate(False)
        tk.Frame(self, bg=C["border"], height=1).pack(side="bottom", fill="x")

        left = tk.Frame(self, bg=C["bg"])
        left.pack(side="left", padx=16, fill="y")

        tk.Label(left, text="Sales Analyzer", bg=C["bg"],
                 fg=C["accent"], font=(UI_FONT, 14, "bold")).pack(
            side="left", pady=0)

        self._file_lbl = tk.Label(left, text="", bg=C["bg"],
                                   fg=C["text_secondary"],
                                   font=(UI_FONT, 11))
        self._file_lbl.pack(side="left", padx=(10, 0))

        right = tk.Frame(self, bg=C["bg"])
        right.pack(side="right", padx=16, fill="y")

        # Version badge
        ver_bg = C["surface"]
        ver = tk.Label(right, text=f" v{APP_VERSION} ",
                       bg=ver_bg, fg=C["text_secondary"],
                       font=(UI_FONT, 9), padx=6, pady=2)
        ver.pack(side="right", padx=(8, 0))

        # Load button (always visible)
        load = _Btn(right, "  📂  Carica  ", on_load,
                    bg=C["surface"], fg=C["text_primary"],
                    font_args=(UI_FONT, 10, "bold"), padx=14, pady=5)
        load.pack(side="right")

        self._update_slot: Optional[tk.Widget] = None

    def set_filename(self, name: str) -> None:
        sep = "  ·  " if name else ""
        self._file_lbl.config(text=f"{sep}{name}")

    def show_update_btn(self, version: str, on_download) -> None:
        """Replace / create the update CTA next to version badge."""
        if self._update_slot and self._update_slot.winfo_exists():
            return
        btn = _Btn(
            self, f"  ↑ v{version}  ", on_download,
            bg=C["accent"], fg="#ffffff",
            font_args=(UI_FONT, 9, "bold"), padx=10, pady=4,
        )
        btn.pack(side="right", padx=(0, 8))
        self._update_slot = btn


# ═══════════════════════════════════════════════════════════════════════════════
# Sidebar
# ═══════════════════════════════════════════════════════════════════════════════

_NAV_ITEM_H = 52   # px per nav item
_NAV_TOP    = 16   # y offset before first item inside sidebar

class Sidebar(tk.Frame):
    """
    200-px sidebar with animated sliding active indicator.
    The indicator bar moves via quintic ease-out over 200 ms.
    """

    _ITEMS = [
        ("trend",      "📈", "Trend"),
        ("top",        "🏆", "Top Prodotti"),
        ("comparison", "📊", "Confronto"),
        ("geo",        "🗺",  "Mappa"),
    ]

    def __init__(self, parent, on_select, **kw):
        super().__init__(parent, bg=C["bg"], width=200, **kw)
        self.pack_propagate(False)
        tk.Frame(self, bg=C["border"], width=1).pack(side="right", fill="y")

        self._on_select = on_select
        self._btns: Dict[str, tk.Frame] = {}
        self._current = "trend"
        self._ind_y: float = _NAV_TOP   # current animated y of indicator

        # "VISTE" header
        tk.Label(self, text="VISTE", bg=C["bg"],
                 fg=C["text_secondary"], font=(UI_FONT, 9, "bold")).pack(
            anchor="w", padx=20, pady=(20, 6))

        # Sliding indicator — positioned with place() over the item area
        self._indicator = tk.Frame(self, bg=C["accent"], width=3)
        self._indicator.place(x=0, y=_NAV_TOP, width=3, height=_NAV_ITEM_H)

        # Nav items
        self._nav_frame = tk.Frame(self, bg=C["bg"])
        self._nav_frame.pack(fill="x")
        for key, icon, label in self._ITEMS:
            row = tk.Frame(self._nav_frame, bg=C["bg"], height=_NAV_ITEM_H)
            row.pack(fill="x")
            row.pack_propagate(False)
            inner = tk.Frame(row, bg=C["bg"])
            inner.place(relx=0, rely=0, relwidth=1, relheight=1)
            tk.Label(inner, text=f"  {icon}  {label}", bg=C["bg"],
                     fg=C["text_secondary"], font=(UI_FONT, 11),
                     anchor="w").place(relx=0, rely=0, relwidth=1, relheight=1)
            inner.bind("<Button-1>",  lambda _, k=key: self._click(k))
            inner.bind("<Enter>",     lambda _, w=inner: w.config(bg=C["surface_elevated"]))
            inner.bind("<Leave>",     lambda _, w=inner: self._restore_row(w))
            for child in inner.winfo_children():
                child.bind("<Button-1>", lambda _, k=key: self._click(k))
                child.bind("<Enter>",    lambda _, w=inner: w.config(bg=C["surface_elevated"]))
                child.bind("<Leave>",    lambda _, w=inner: self._restore_row(w))
            self._btns[key] = inner

        self._set_active_style("trend")

    def _idx(self, key: str) -> int:
        return [k for k, _, _ in self._ITEMS].index(key)

    def _click(self, key: str) -> None:
        if key == self._current:
            return
        prev = self._current
        self._current = key
        self._set_active_style(prev)
        self._set_active_style(key)
        self._slide_indicator(key)
        self._on_select(key)

    def _set_active_style(self, key: str) -> None:
        active = (key == self._current)
        w = self._btns[key]
        bg = C["surface"] if active else C["bg"]
        fg = C["accent"] if active else C["text_secondary"]
        w.config(bg=bg)
        for child in w.winfo_children():
            child.config(bg=bg, fg=fg)

    def _restore_row(self, w: tk.Frame) -> None:
        key = next((k for k, v in self._btns.items() if v is w), None)
        if key:
            self._set_active_style(key)

    def _slide_indicator(self, key: str) -> None:
        """Slide the 3-px accent bar to the new active item (200 ms, quintic)."""
        target_y = float(_NAV_TOP + self._idx(key) * _NAV_ITEM_H)
        start_y  = self._ind_y

        def _step(te: float) -> None:
            y = start_y + (target_y - start_y) * te
            self._ind_y = y
            self._indicator.place(x=0, y=int(y), width=3,
                                  height=_NAV_ITEM_H)

        _run_anim(self, 200, _ease_quint, _step)

    def set_active(self, key: str) -> None:
        """External call to set active item without firing on_select."""
        if key == self._current:
            return
        prev = self._current
        self._current = key
        self._set_active_style(prev)
        self._set_active_style(key)
        self._slide_indicator(key)


# ═══════════════════════════════════════════════════════════════════════════════
# Filter bar
# ═══════════════════════════════════════════════════════════════════════════════

class FilterBar(tk.Frame):
    """Compact filter strip: ⚙ Colonne | dates | channel | Applica | Reset · 📊 Grafico."""

    # Dimmed fg color used when buttons are disabled (pre-file-load)
    _DIM_FG = "#3a3a3a"

    def __init__(self, parent, on_apply, on_reset,
                 on_col_config=None, on_chart_type=None, **kw):
        super().__init__(parent, bg=C["surface"], height=48, **kw)
        self.pack_propagate(False)
        tk.Frame(self, bg=C["border"], height=1).pack(side="top",    fill="x")
        tk.Frame(self, bg=C["border"], height=1).pack(side="bottom", fill="x")
        self._on_col_config  = on_col_config
        self._on_chart_type  = on_chart_type
        self._enabled        = False
        self._col_badge      = False

        row = tk.Frame(self, bg=C["surface"])
        row.pack(fill="both", expand=True, padx=8)

        def sep():
            tk.Frame(row, bg=C["border"], width=1).pack(
                side="left", fill="y", pady=8, padx=4)

        # ── LEFT: ⚙ Colonne ──
        self._col_btn = tk.Label(
            row, text="⚙  Colonne", bg=C["surface"], fg=self._DIM_FG,
            font=(UI_FONT, 10), padx=10, pady=5)
        self._col_btn.pack(side="left", padx=(2, 0))
        self._col_btn.bind("<Button-1>", lambda _: self._col_clicked())

        sep()

        def lbl(text):
            return tk.Label(row, text=text, bg=C["surface"],
                            fg=C["text_secondary"], font=(UI_FONT, 10))

        lbl("Da").pack(side="left", padx=(0, 4))
        self._date_from = self._date_widget(row)
        self._date_from.pack(side="left", padx=(0, 6))

        lbl("→").pack(side="left", padx=4)

        lbl("A").pack(side="left", padx=(6, 4))
        self._date_to = self._date_widget(row)
        self._date_to.pack(side="left", padx=(0, 8))

        sep()

        self._ch_dd = ChannelDropdown(row, bg=C["surface"])
        self._ch_dd.pack(side="left", padx=6)

        sep()

        apply_btn = _Btn(row, "  Applica  ", on_apply,
                         bg=C["accent"], fg="#ffffff",
                         font_args=(UI_FONT, 10, "bold"))
        apply_btn.pack(side="left", padx=4)

        reset_lbl = tk.Label(row, text="Reset", bg=C["surface"],
                             fg=C["text_secondary"], font=(UI_FONT, 10),
                             cursor="hand2")
        reset_lbl.pack(side="left", padx=(2, 0))
        reset_lbl.bind("<Button-1>", lambda _: on_reset())

        # ── RIGHT: 📊 Grafico ──
        self._chart_btn = tk.Label(
            row, text="📊  Grafico", bg=C["surface"], fg=self._DIM_FG,
            font=(UI_FONT, 10), padx=10, pady=5)
        self._chart_btn.pack(side="right", padx=(0, 4))
        self._chart_btn.bind("<Button-1>", lambda _: self._chart_clicked())

    # ── Button state helpers ──────────────────────────────────────────────────

    def set_buttons_enabled(self, enabled: bool) -> None:
        self._enabled = enabled
        fg = C["text_secondary"] if enabled else self._DIM_FG
        cur = "hand2" if enabled else ""
        self._col_btn.config(fg=fg, cursor=cur)
        self._chart_btn.config(fg=fg, cursor=cur)
        self._redraw_col_badge()

    def set_col_badge(self, active: bool) -> None:
        self._col_badge = active
        self._redraw_col_badge()

    def _redraw_col_badge(self) -> None:
        if self._col_badge and self._enabled:
            self._col_btn.config(
                text="⚙  Colonne  ●",
                fg="#f97316")   # orange dot
        else:
            self._col_btn.config(
                text="⚙  Colonne",
                fg=C["text_secondary"] if self._enabled else self._DIM_FG)

    def _col_clicked(self) -> None:
        if self._enabled and self._on_col_config:
            self._on_col_config()

    def _chart_clicked(self) -> None:
        if self._enabled and self._on_chart_type:
            self._on_chart_type(self._chart_btn)

    @property
    def col_btn(self): return self._col_btn

    @property
    def chart_btn(self): return self._chart_btn

    def _date_widget(self, parent) -> tk.Widget:
        if HAS_CALENDAR:
            de = DateEntry(parent, width=10, date_pattern="dd/MM/yyyy",
                           background=C["surface_elevated"],
                           foreground=C["text_primary"],
                           bordercolor=C["border"],
                           headersbackground=C["surface"],
                           headersforeground=C["text_primary"],
                           selectbackground=C["accent"],
                           normalbackground=C["surface"],
                           normalforeground=C["text_primary"],
                           weekendbackground=C["surface"],
                           weekendforeground=C["text_secondary"],
                           font=(UI_FONT, 9))
            return de
        e = tk.Entry(parent, width=11, bg=C["surface_elevated"],
                     fg=C["text_primary"], insertbackground=C["text_primary"],
                     relief="flat", font=(UI_FONT, 9),
                     highlightthickness=1, highlightcolor=C["border"],
                     highlightbackground=C["border"])
        e.insert(0, "GG/MM/AAAA")
        e.bind("<FocusIn>",
               lambda ev, w=e: w.delete(0, "end")
               if w.get() == "GG/MM/AAAA" else None)
        return e

    def get_date(self, widget) -> Optional[date]:
        if HAS_CALENDAR:
            try:
                return widget.get_date()
            except Exception:
                return None
        val = widget.get().strip()
        if not val or val == "GG/MM/AAAA":
            return None
        for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d"):
            try:
                return datetime.strptime(val, fmt).date()
            except ValueError:
                pass
        return None

    def set_date(self, widget, d: date) -> None:
        if HAS_CALENDAR:
            try:
                widget.set_date(d)
            except Exception:
                pass
        else:
            widget.delete(0, "end")
            widget.insert(0, d.strftime("%d/%m/%Y"))

    @property
    def date_from(self): return self._date_from
    @property
    def date_to(self):   return self._date_to
    @property
    def channel_dd(self): return self._ch_dd


# ═══════════════════════════════════════════════════════════════════════════════
# Column Mapping Dialog
# ═══════════════════════════════════════════════════════════════════════════════

class ColumnMappingDialog(tk.Toplevel):
    _ROLES = {
        "date":     ("Data / Periodo *", True),
        "product":  ("Prodotto / Categoria *", True),
        "revenue":  ("Ricavo / Importo *", True),
        "quantity": ("Quantità Venduta", False),
        "channel":  ("Canale di Vendita", False),
        "geo":      ("Città / Regione", False),
    }

    def __init__(self, parent, columns: List[str],
                 mapping: Dict[str, Optional[str]]):
        super().__init__(parent)
        self.title("Mappa Colonne Excel")
        self.configure(bg=C["bg"])
        self.resizable(False, False)
        self.grab_set()
        self.result: Optional[Dict] = None
        self._combos: Dict[str, ttk.Combobox] = {}

        choices = ["— Non presente —"] + list(columns)

        tk.Label(self, text="Associa le colonne Excel ai campi richiesti:",
                 bg=C["bg"], fg=C["text_primary"],
                 font=(UI_FONT, 10)).pack(pady=(18, 8), padx=24)

        form = tk.Frame(self, bg=C["bg"])
        form.pack(padx=24, pady=4)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("M.TCombobox",
                        fieldbackground=C["surface_elevated"],
                        background=C["surface_elevated"],
                        foreground=C["text_primary"],
                        selectbackground=C["accent"],
                        bordercolor=C["border"])
        style.map("M.TCombobox",
                  fieldbackground=[("readonly", C["surface_elevated"])],
                  foreground=[("readonly", C["text_primary"])])

        for row_i, (role, (label, required)) in enumerate(self._ROLES.items()):
            txt = label if required else label + " (opzionale)"
            tk.Label(form, text=txt, bg=C["bg"], fg=C["text_primary"],
                     font=(UI_FONT, 9), anchor="w",
                     width=28).grid(row=row_i, column=0, sticky="w",
                                    pady=5, padx=(0, 14))
            cb = ttk.Combobox(form, values=choices, state="readonly",
                               width=28, style="M.TCombobox")
            cur = mapping.get(role)
            cb.set(cur if cur else "— Non presente —")
            cb.grid(row=row_i, column=1, pady=5)
            self._combos[role] = cb

        btns = tk.Frame(self, bg=C["bg"])
        btns.pack(pady=18)
        _Btn(btns, "Annulla", self.destroy,
             bg=C["surface"], padx=18, pady=7).pack(side="left", padx=8)
        _Btn(btns, "Conferma", self._confirm,
             bg=C["accent"], fg="#ffffff", padx=18, pady=7).pack(
            side="left", padx=8)

        self.update_idletasks()
        pw = parent.winfo_width(); ph = parent.winfo_height()
        px = parent.winfo_rootx(); py = parent.winfo_rooty()
        w = self.winfo_reqwidth(); h = self.winfo_reqheight()
        self.geometry(f"+{px+(pw-w)//2}+{py+(ph-h)//2}")

    def _confirm(self) -> None:
        result = {r: (None if cb.get().startswith("—") else cb.get())
                  for r, cb in self._combos.items()}
        missing = [r for r in ("date", "product", "revenue")
                   if not result.get(r)]
        if missing:
            messagebox.showerror(
                "Campi Mancanti",
                "Associa i campi obbligatori:\n" +
                "\n".join(f"  • {self._ROLES[r][0]}" for r in missing),
                parent=self)
            return
        self.result = result
        self.destroy()


# ═══════════════════════════════════════════════════════════════════════════════
# Download Dialog
# ═══════════════════════════════════════════════════════════════════════════════

class DownloadDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Aggiornamento")
        self.configure(bg=C["bg"])
        self.resizable(False, False)
        self.grab_set()
        self.geometry("440x160")
        self._center(parent)

        tk.Label(self, text="Download aggiornamento…",
                 bg=C["bg"], fg=C["text_primary"],
                 font=(UI_FONT, 10)).pack(pady=(22, 8))

        style = ttk.Style()
        style.configure("D.Horizontal.TProgressbar",
                        troughcolor=C["surface"],
                        background=C["accent"],
                        bordercolor=C["border"])
        self._pb = ttk.Progressbar(self, style="D.Horizontal.TProgressbar",
                                    length=400, mode="determinate")
        self._pb.pack(pady=4, padx=20)

        self._lbl = tk.Label(self, text="", bg=C["bg"],
                             fg=C["text_secondary"], font=(UI_FONT, 9))
        self._lbl.pack()

    def _center(self, parent) -> None:
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_rootx(), parent.winfo_rooty()
        self.geometry(f"+{px+(pw-440)//2}+{py+(ph-160)//2}")

    def update_progress(self, downloaded: int, total: int) -> None:
        def _do():
            mb_d = downloaded / 1_048_576
            mb_t = total / 1_048_576 if total else 0
            self._lbl.config(text=f"{mb_d:.1f} / {mb_t:.1f} MB")
            self._pb["value"] = downloaded / total * 100 if total else 0
        self.after(0, _do)

    def show_error(self, msg: str, on_retry) -> None:
        def _do():
            self._lbl.config(text=f"Errore: {msg}", fg=C["accent_warm"])
            _Btn(self, "Riprova", on_retry,
                 bg=C["accent"], fg="#ffffff", padx=12, pady=4).pack(pady=8)
        self.after(0, _do)

    def show_complete(self) -> None:
        def _do():
            self._pb["value"] = 100
            self._lbl.config(
                text="Scaricato. Riavvia l'app per applicare.",
                fg=C["accent_green"])
        self.after(0, _do)


# ═══════════════════════════════════════════════════════════════════════════════
# Base chart view
# ═══════════════════════════════════════════════════════════════════════════════

class BaseView(tk.Frame):
    VIEW_TITLE    = "Vista"
    VIEW_KEY      = ""       # set by subclasses; matches CHART_COMPAT keys
    DEFAULT_CHART = "line"   # override in subclasses

    def __init__(self, parent, app: "SalesAnalyzerApp"):
        super().__init__(parent, bg=C["bg"])
        self.app = app
        self._fig:        Optional[Figure]            = None
        self._canvas:     Optional[FigureCanvasTkAgg] = None
        self._chart_type: str = self.DEFAULT_CHART

        # Per-view controls bar (view subclasses populate this)
        self._ctrl = tk.Frame(self, bg=C["bg"])
        self._ctrl.pack(fill="x", padx=12, pady=(10, 0))

        # Chart frame
        self._chart = tk.Frame(self, bg=C["chart_bg"])
        self._chart.pack(fill="both", expand=True, padx=12, pady=8)

    # ── Matplotlib helpers ────────────────────────────────────────────────────

    def _make_fig(self, **kw) -> Tuple[Figure, Any]:
        """Clear previous canvas, create a new figure/axes pair."""
        if self._canvas:
            self._canvas.get_tk_widget().destroy()
            self._canvas = None
        fig = Figure(facecolor=C["chart_bg"], **kw)
        self._fig = fig
        cv  = FigureCanvasTkAgg(fig, master=self._chart)
        cv.get_tk_widget().pack(fill="both", expand=True)
        self._canvas = cv
        ax = fig.add_subplot(1, 1, 1)
        return fig, ax

    def _style(self, ax) -> None:
        """Apply shared dark-theme styling to an axes object."""
        ax.set_facecolor(C["chart_bg"])
        ax.tick_params(colors=C["text_secondary"], labelsize=10,
                       labelcolor=C["text_secondary"])
        ax.grid(True, color=C["chart_grid"], linewidth=0.5,
                linestyle="--", alpha=0.8)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color(C["border"])
        ax.spines["bottom"].set_color(C["border"])
        ax.title.set_color(C["text_primary"])
        ax.title.set_fontsize(13)
        ax.title.set_fontfamily(UI_FONT)
        ax.title.set_fontweight("bold")
        ax.xaxis.label.set_color(C["text_secondary"])
        ax.xaxis.label.set_fontsize(11)
        ax.yaxis.label.set_color(C["text_secondary"])
        ax.yaxis.label.set_fontsize(11)

    def _clear(self) -> None:
        """Destroy all chart-frame children (canvas, empty-state labels, etc.)."""
        if self._canvas:
            self._canvas.get_tk_widget().destroy()
            self._canvas = None
        self._fig = None
        for w in self._chart.winfo_children():
            w.destroy()

    def _empty(self, msg: str = "Nessun dato") -> None:
        self._clear()
        outer = tk.Frame(self._chart, bg=C["chart_bg"])
        outer.pack(expand=True)
        tk.Label(outer, text=msg, bg=C["chart_bg"],
                 fg=C["text_secondary"], font=(UI_FONT, 12),
                 wraplength=360, justify="center").pack(pady=60)

    def _draw_done(self) -> None:
        """Call after canvas.draw() to trigger the wipe-reveal animation."""
        if self._canvas:
            self._canvas.draw()
            _wipe_reveal(self._chart, 380)

    def export_png(self) -> None:
        if not self._fig:
            messagebox.showinfo("Export", "Nessun grafico disponibile.", parent=self)
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG", "*.png")],
            initialfile=f"{self.VIEW_TITLE}.png")
        if path:
            self._fig.savefig(path, dpi=150, bbox_inches="tight",
                              facecolor=C["chart_bg"])
            messagebox.showinfo("Export", f"Salvato:\n{path}", parent=self)

    def set_chart_type(self, ct: str) -> None:
        self._chart_type = ct
        self.refresh_current()

    def refresh_current(self) -> None:
        pass

    def refresh(self, df: pd.DataFrame, mapping: Dict) -> None:
        raise NotImplementedError

    # ── Hover tooltip helpers ─────────────────────────────────────────────────

    def _make_annot(self, ax):
        """Styled hover annotation matching the dark theme."""
        ann = ax.annotate(
            "", xy=(0, 0), xytext=(14, 14),
            textcoords="offset points",
            bbox=dict(boxstyle="round,pad=0.55",
                      facecolor=C["surface_elevated"],
                      edgecolor=C["border"],
                      linewidth=1, alpha=0.96),
            fontsize=10, color=C["text_primary"],
            fontfamily=MONO_FONT, zorder=10,
        )
        ann.set_visible(False)
        return ann

    def _hover_line(self, ax, x_labels: list, ys, extra: Optional[Dict] = None):
        """
        Attach hover to a line/area chart.
        Snaps to the nearest plotted point on ax.lines[0].
        extra: optional {label: ys_array} for additional series (e.g. Qty axis).
        """
        annot  = self._make_annot(ax)
        canvas = self._canvas
        ys_f   = np.asarray(ys, dtype=float)
        n      = len(ys_f)

        def on_move(event):
            if event.inaxes is not ax:
                if annot.get_visible():
                    annot.set_visible(False); canvas.draw_idle()
                return
            xd = event.xdata
            if xd is None or not ax.lines:
                return
            # Compare against the line's own x data (date floats or ints)
            line_xs = ax.lines[0].get_xdata()
            if len(line_xs) == 0:
                return
            idx = int(np.argmin(np.abs(np.asarray(line_xs, dtype=float) - xd)))
            idx = max(0, min(n - 1, idx))
            xv, yv = line_xs[idx], ys_f[idx]
            parts = [str(x_labels[idx]), f"€{yv:,.0f}"]
            if extra:
                for name, ey in extra.items():
                    ef = np.asarray(ey, dtype=float)
                    if idx < len(ef):
                        parts.append(f"{name}: {ef[idx]:,.0f}")
            annot.set_text("\n".join(parts))
            annot.xy = (xv, yv)
            xlim = ax.get_xlim(); ylim = ax.get_ylim()
            xf = (float(xv) - xlim[0]) / max(xlim[1] - xlim[0], 1e-9)
            yf = (yv - ylim[0]) / max(ylim[1] - ylim[0], 1e-9)
            annot.xytext = (-100 if xf > 0.75 else 14,
                            -40  if yf > 0.80 else 14)
            annot.set_visible(True); canvas.draw_idle()

        canvas.mpl_connect("motion_notify_event", on_move)

    def _hover_multi_lines(self, ax, x_labels: list, year_vals: dict):
        """Hover for YoY multi-line chart. Shows all years' values at nearest month."""
        annot  = self._make_annot(ax)
        canvas = self._canvas
        n      = len(x_labels)

        def on_move(event):
            if event.inaxes is not ax:
                if annot.get_visible():
                    annot.set_visible(False); canvas.draw_idle()
                return
            xd = event.xdata
            if xd is None:
                return
            idx = max(0, min(n - 1, int(round(xd))))
            parts = [x_labels[idx]]
            for yr, vals in year_vals.items():
                if idx < len(vals):
                    parts.append(f"{yr}: €{float(vals[idx]):,.0f}")
            annot.set_text("\n".join(parts))
            annot.xy = (idx, event.ydata or 0)
            xlim = ax.get_xlim()
            xf = (idx - xlim[0]) / max(xlim[1] - xlim[0], 1e-9)
            annot.xytext = (-120 if xf > 0.72 else 14, 14)
            annot.set_visible(True); canvas.draw_idle()

        canvas.mpl_connect("motion_notify_event", on_move)

    def _hover_bars_v(self, ax, all_bars: list, all_labels: list, all_vals: list):
        """
        Hover for vertical bar charts.
        all_bars/all_labels/all_vals are flat lists with one entry per bar.
        """
        annot   = self._make_annot(ax)
        canvas  = self._canvas
        triples = list(zip(all_bars, all_labels, all_vals))

        def on_move(event):
            if event.inaxes is not ax:
                if annot.get_visible():
                    annot.set_visible(False); canvas.draw_idle()
                return
            for b, lbl, val in triples:
                if b.contains(event)[0]:
                    cx = b.get_x() + b.get_width() / 2
                    cy = b.get_height()
                    annot.set_text(f"{lbl}\n€{float(val):,.0f}")
                    annot.xy = (cx, cy)
                    xlim = ax.get_xlim(); ylim = ax.get_ylim()
                    xf = (cx - xlim[0]) / max(xlim[1] - xlim[0], 1e-9)
                    yf = (cy - ylim[0]) / max(ylim[1] - ylim[0], 1e-9)
                    annot.xytext = (-100 if xf > 0.75 else 14,
                                    -40  if yf > 0.85 else 8)
                    annot.set_visible(True); canvas.draw_idle()
                    return
            if annot.get_visible():
                annot.set_visible(False); canvas.draw_idle()

        canvas.mpl_connect("motion_notify_event", on_move)

    def _hover_bars_h(self, ax, bars, labels: list, vals):
        """Hover for horizontal bar charts."""
        annot   = self._make_annot(ax)
        canvas  = self._canvas
        triples = list(zip(bars, labels, vals))

        def on_move(event):
            if event.inaxes is not ax:
                if annot.get_visible():
                    annot.set_visible(False); canvas.draw_idle()
                return
            for b, lbl, val in triples:
                if b.contains(event)[0]:
                    cx = b.get_width()
                    cy = b.get_y() + b.get_height() / 2
                    annot.set_text(f"{lbl}\n€{float(val):,.0f}")
                    annot.xy = (cx, cy)
                    xlim = ax.get_xlim()
                    xf = (cx - xlim[0]) / max(xlim[1] - xlim[0], 1e-9)
                    annot.xytext = (-120 if xf > 0.75 else 12, -8)
                    annot.set_visible(True); canvas.draw_idle()
                    return
            if annot.get_visible():
                annot.set_visible(False); canvas.draw_idle()

        canvas.mpl_connect("motion_notify_event", on_move)

    def _hover_pie(self, ax, wedges, labels: list, vals):
        """Hover for pie/donut charts."""
        annot  = self._make_annot(ax)
        canvas = self._canvas
        total  = sum(float(v) for v in vals)
        quads  = list(zip(wedges, labels, vals))

        def on_move(event):
            if event.inaxes is not ax:
                if annot.get_visible():
                    annot.set_visible(False); canvas.draw_idle()
                return
            for w, lbl, val in quads:
                if w.contains(event)[0]:
                    pct = float(val) / total * 100 if total else 0
                    annot.set_text(f"{lbl}\n€{float(val):,.0f}\n{pct:.1f}%")
                    annot.xy = (event.xdata, event.ydata)
                    annot.xytext = (14, 14)
                    annot.set_visible(True); canvas.draw_idle()
                    return
            if annot.get_visible():
                annot.set_visible(False); canvas.draw_idle()

        canvas.mpl_connect("motion_notify_event", on_move)


# ═══════════════════════════════════════════════════════════════════════════════
# View 1 — Trend
# ═══════════════════════════════════════════════════════════════════════════════

class TrendView(BaseView):
    VIEW_TITLE    = "Trend Ricavi"
    VIEW_KEY      = "trend"
    DEFAULT_CHART = "line"

    def __init__(self, parent, app):
        super().__init__(parent, app)
        self._show_qty = tk.BooleanVar(value=False)
        self._df: Optional[pd.DataFrame] = None
        self._mapping: Optional[Dict] = None

        # Pill toggle for granularity
        self._gran = PillGroup(
            self._ctrl,
            ["Giorno", "Settimana", "Mese", "Trimestre"],
            initial="Mese",
            callback=self._on_gran,
            bg=C["bg"],
        )
        self._gran.pack(side="left")

        tk.Checkbutton(
            self._ctrl, text="Mostra Quantità",
            variable=self._show_qty,
            bg=C["bg"], fg=C["text_secondary"],
            selectcolor=C["accent"],
            activebackground=C["bg"], activeforeground=C["text_primary"],
            font=(UI_FONT, 10), cursor="hand2",
            command=self._redraw,
        ).pack(side="left", padx=16)

    def _on_gran(self, _: str) -> None:
        self._redraw()

    def _redraw(self) -> None:
        if self._df is not None:
            self._draw(self._df, self._mapping)

    def refresh_current(self) -> None:
        self._redraw()

    def refresh(self, df: pd.DataFrame, mapping: Dict) -> None:
        self._df      = df
        self._mapping = mapping
        self._draw(df, mapping)

    def _draw(self, df: pd.DataFrame, mapping: Dict) -> None:
        self._clear()
        if df is None or df.empty:
            self._empty("Nessun dato da visualizzare.")
            return

        date_col = mapping["date"]
        rev_col  = mapping["revenue"]
        qty_col  = mapping.get("quantity")

        freq_map = {"Giorno":"D","Settimana":"W","Mese":"ME","Trimestre":"QE"}
        freq = freq_map[self._gran.get()]

        ts = df.set_index(date_col)[[rev_col]].copy()
        if qty_col and self._show_qty.get():
            ts[qty_col] = df.set_index(date_col)[qty_col]

        try:
            grouped = ts.resample(freq).sum()
        except Exception:
            grouped = ts.resample({"ME":"M","QE":"Q"}.get(freq, freq)).sum()

        if grouped.empty:
            self._empty("Nessun dato per il periodo selezionato.")
            return

        ct  = self._chart_type
        fig, ax = self._make_fig(figsize=(10, 5))
        self._style(ax)

        title = f"Andamento Ricavi — {self._gran.get()}"
        ax.set_title(title, pad=14)
        ax.set_ylabel("Ricavi (€)")
        ax.yaxis.set_major_formatter(
            mticker.FuncFormatter(lambda x, _: f"€{x:,.0f}"))

        x_labels = [str(grouped.index[i])[:10] for i in range(len(grouped))]

        if ct == "bar_v":
            xs = range(len(grouped))
            colors = _bar_gradient(len(grouped))
            trend_bars = ax.bar(xs, grouped[rev_col].values, color=colors, width=0.7)
            step = max(1, len(grouped) // 10)
            ax.set_xticks(list(xs)[::step])
            ax.set_xticklabels(
                [x_labels[i] for i in range(0, len(grouped), step)],
                rotation=40, ha="right", color=C["text_secondary"], fontsize=8)
            fig.tight_layout()
            self._hover_bars_v(ax,
                               list(trend_bars),
                               x_labels,
                               grouped[rev_col].values.tolist())
        elif ct == "area":
            ax.fill_between(grouped.index, grouped[rev_col],
                            alpha=0.35, color=C["accent"])
            ax.plot(grouped.index, grouped[rev_col],
                    color=C["accent"], linewidth=0)
            fig.tight_layout()
            self._hover_line(ax, x_labels, grouped[rev_col].values)
        else:  # "line" (default)
            ax.plot(grouped.index, grouped[rev_col],
                    color=C["accent"], linewidth=2.5,
                    marker="o", markersize=3, label="Ricavi")
            ax.fill_between(grouped.index, grouped[rev_col],
                            alpha=0.15, color=C["accent"])

            qty_extra = None
            if qty_col and self._show_qty.get() and qty_col in grouped.columns:
                ax2 = ax.twinx()
                ax2.set_facecolor(C["chart_bg"])
                ax2.tick_params(colors=C["text_secondary"], labelsize=10)
                for sp in ("top","left"):
                    ax2.spines[sp].set_visible(False)
                ax2.spines["right"].set_color(C["border"])
                ax2.spines["bottom"].set_color(C["border"])
                ax2.plot(grouped.index, grouped[qty_col],
                         color=C["accent_green"], linewidth=2,
                         linestyle="--", marker="s", markersize=3,
                         label="Quantità")
                ax2.set_ylabel("Quantità", color=C["text_secondary"], fontsize=11)
                h1, l1 = ax.get_legend_handles_labels()
                h2, l2 = ax2.get_legend_handles_labels()
                ax.legend(h1+h2, l1+l2, framealpha=0,
                          labelcolor=C["text_secondary"], fontsize=10)
                qty_extra = {"Qtà": grouped[qty_col].values}
            else:
                ax.legend(framealpha=0, labelcolor=C["text_secondary"], fontsize=10)

            fig.tight_layout()
            self._hover_line(ax, x_labels, grouped[rev_col].values,
                             extra=qty_extra)

        self._draw_done()


# ═══════════════════════════════════════════════════════════════════════════════
# View 2 — Top Products
# ═══════════════════════════════════════════════════════════════════════════════

class TopProductsView(BaseView):
    VIEW_TITLE    = "Top Prodotti"
    VIEW_KEY      = "top"
    DEFAULT_CHART = "bar_h"

    def __init__(self, parent, app):
        super().__init__(parent, app)
        self._chart_type = "bar_h"
        self._df: Optional[pd.DataFrame] = None
        self._mapping: Optional[Dict] = None

        self._n_pill = PillGroup(
            self._ctrl, ["5","10","20","Tutti"],
            initial="10", callback=self._on_n, bg=C["bg"])
        self._n_pill.pack(side="left")

    def _on_n(self, _: str) -> None:
        if self._df is not None:
            self._draw(self._df, self._mapping)

    def refresh_current(self) -> None:
        if self._df is not None:
            self._draw(self._df, self._mapping)

    def refresh(self, df: pd.DataFrame, mapping: Dict) -> None:
        self._df      = df
        self._mapping = mapping
        self._draw(df, mapping)

    def _draw(self, df: pd.DataFrame, mapping: Dict) -> None:
        self._clear()
        if df is None or df.empty:
            self._empty("Nessun dato da visualizzare.")
            return

        prod_col = mapping["product"]
        rev_col  = mapping["revenue"]

        agg = df.groupby(prod_col)[rev_col].sum().sort_values(ascending=False)
        n_str = self._n_pill.get()
        if n_str != "Tutti":
            agg = agg.head(int(n_str))
        if agg.empty:
            self._empty("Nessun prodotto trovato.")
            return

        ct = self._chart_type
        n  = len(agg)

        if ct == "pie":
            self._draw_pie(agg, n_str)
            return
        if ct == "bar_v":
            self._draw_bar_v(agg, n_str, n)
            return
        # default: bar_h
        self._draw_bar_h(agg, n_str, n)

    def _draw_bar_h(self, agg, n_str, n) -> None:
        fig_h = max(4.0, min(n * 0.48 + 1.2, 12.0))
        fig, ax = self._make_fig(figsize=(10, fig_h))
        self._style(ax)
        ax.grid(axis="x", color=C["chart_grid"], linewidth=0.5,
                linestyle="--", alpha=0.8)
        ax.grid(axis="y", visible=False)

        colors = _bar_gradient(n)
        bars = ax.barh(range(n), agg.values, color=colors, height=0.6)
        ax.set_yticks(range(n))
        ax.set_yticklabels([str(p)[:32] for p in agg.index],
                           fontsize=10, color=C["text_primary"])
        ax.invert_yaxis()

        mx = agg.max()
        for bar, val in zip(bars, agg.values):
            ax.text(bar.get_width() + mx * 0.012,
                    bar.get_y() + bar.get_height() / 2,
                    f"€{val:,.0f}",
                    va="center", ha="left",
                    color=C["text_secondary"],
                    fontsize=9, fontfamily=MONO_FONT)

        ax.set_xlabel("Ricavo Totale (€)")
        ax.set_title(f"Top {n_str} Prodotti per Ricavo", pad=14)
        ax.xaxis.set_major_formatter(
            mticker.FuncFormatter(lambda x, _: f"€{x:,.0f}"))
        fig.tight_layout()
        self._hover_bars_h(ax, bars,
                           [str(p) for p in agg.index],
                           agg.values.tolist())
        self._draw_done()

    def _draw_bar_v(self, agg, n_str, n) -> None:
        fig, ax = self._make_fig(figsize=(10, 5))
        self._style(ax)
        ax.grid(axis="y", color=C["chart_grid"], linewidth=0.5,
                linestyle="--", alpha=0.8)
        ax.grid(axis="x", visible=False)

        colors = _bar_gradient(n)
        xs = range(n)
        top_bars = ax.bar(xs, agg.values, color=colors, width=0.6)
        ax.set_xticks(list(xs))
        ax.set_xticklabels([str(p)[:20] for p in agg.index],
                           rotation=45, ha="right",
                           fontsize=9, color=C["text_primary"])
        ax.set_ylabel("Ricavo Totale (€)")
        ax.set_title(f"Top {n_str} Prodotti per Ricavo", pad=14)
        ax.yaxis.set_major_formatter(
            mticker.FuncFormatter(lambda x, _: f"€{x:,.0f}"))
        fig.tight_layout()
        self._hover_bars_v(ax,
                           list(top_bars),
                           [str(p) for p in agg.index],
                           agg.values.tolist())
        self._draw_done()

    def _draw_pie(self, agg, n_str) -> None:
        fig, ax = self._make_fig(figsize=(9, 6))
        ax.set_facecolor(C["chart_bg"])
        ax.set_title(f"Top {n_str} Prodotti — Quota di Ricavo", pad=14,
                     color=C["text_primary"], fontsize=13,
                     fontfamily=UI_FONT, fontweight="bold")

        colors = CHART_COLORS[:len(agg)]
        wedges, texts, autotexts = ax.pie(
            agg.values,
            labels=None,
            colors=colors,
            autopct="%1.1f%%",
            pctdistance=0.75,
            startangle=90,
            wedgeprops=dict(width=0.5, edgecolor=C["chart_bg"], linewidth=2),
        )
        for t in autotexts:
            t.set_color(C["text_primary"])
            t.set_fontsize(8)

        ax.legend(wedges, [str(p)[:24] for p in agg.index],
                  loc="center left", bbox_to_anchor=(1, 0.5),
                  framealpha=0, labelcolor=C["text_secondary"], fontsize=9)
        self._fig.tight_layout()
        self._hover_pie(ax, wedges,
                        [str(p) for p in agg.index],
                        agg.values.tolist())
        self._draw_done()


# ═══════════════════════════════════════════════════════════════════════════════
# View 3 — Period Comparison
# ═══════════════════════════════════════════════════════════════════════════════

class PeriodComparisonView(BaseView):
    VIEW_TITLE    = "Confronto Periodi"
    VIEW_KEY      = "comparison"
    DEFAULT_CHART = "bar_v"

    def __init__(self, parent, app):
        super().__init__(parent, app)
        self._chart_type = "bar_v"
        self._df: Optional[pd.DataFrame] = None
        self._mapping: Optional[Dict] = None
        self._stats_bar: Optional[tk.Frame] = None

        self._mode_pill = PillGroup(
            self._ctrl, ["YoY","MoM"],
            initial="YoY", callback=self._on_mode, bg=C["bg"])
        self._mode_pill.pack(side="left")

    def _on_mode(self, _: str) -> None:
        if self._df is not None:
            self._draw_dispatch()

    def refresh_current(self) -> None:
        if self._df is not None:
            self._draw_dispatch()

    def refresh(self, df: pd.DataFrame, mapping: Dict) -> None:
        self._df      = df
        self._mapping = mapping
        self._draw_dispatch()

    def _draw_dispatch(self) -> None:
        if self._stats_bar:
            self._stats_bar.destroy()
            self._stats_bar = None
        self._clear()
        if self._df is None or self._df.empty:
            self._empty("Nessun dato da visualizzare.")
            return
        dc = self._mapping["date"]
        rc = self._mapping["revenue"]
        mode = self._mode_pill.get()
        ct   = self._chart_type
        if mode == "YoY":
            if ct == "line":
                self._draw_yoy_line(self._df, dc, rc)
            else:
                self._draw_yoy(self._df, dc, rc)
        else:
            if ct == "line":
                self._draw_mom_line(self._df, dc, rc)
            else:
                self._draw_mom(self._df, dc, rc)

    def _draw_yoy(self, df, date_col, rev_col) -> None:
        df = df.copy()
        df["_y"] = df[date_col].dt.year
        df["_m"] = df[date_col].dt.month
        pivot = df.groupby(["_y","_m"])[rev_col].sum().unstack(level=0)
        if pivot.empty:
            self._empty("Dati insufficienti per YoY.")
            return

        years  = sorted(pivot.columns)
        months = list(range(1, 13))
        mnames = ["Gen","Feb","Mar","Apr","Mag","Giu",
                  "Lug","Ago","Set","Ott","Nov","Dic"]

        fig, ax = self._make_fig(figsize=(12, 5))
        self._style(ax)
        ax.grid(axis="y"); ax.grid(axis="x", visible=False)

        ny    = len(years)
        bw    = 0.7 / ny
        xs    = np.arange(len(months))

        yr_colors = [
            C["text_secondary"] if i < ny - 1 else C["accent"]
            for i in range(ny)
        ]

        all_yoy_bars: list = []
        all_yoy_lbls: list = []
        all_yoy_vals: list = []

        for i, year in enumerate(years):
            vals   = [float(pivot.loc[m, year])
                      if m in pivot.index and year in pivot.columns
                      else 0.0 for m in months]
            offset = (i - (ny - 1) / 2) * bw
            alpha  = 0.4 if i < ny - 1 else 1.0
            bars   = ax.bar(xs + offset, vals, bw * 0.9,
                            color=yr_colors[i], alpha=alpha,
                            label=str(year))
            for j, (bar, m) in enumerate(zip(bars, months)):
                all_yoy_bars.append(bar)
                all_yoy_lbls.append(f"{mnames[j]} {year}")
                all_yoy_vals.append(vals[j])
                if i > 0:
                    prev = years[i - 1]
                    pv = float(pivot.loc[m, prev]
                               if m in pivot.index and prev in pivot.columns
                               else 0)
                    cv = vals[j]
                    if pv and cv:
                        pct = (cv - pv) / pv * 100
                        col = C["accent_green"] if pct >= 0 else C["accent_warm"]
                        ax.text(bar.get_x() + bar.get_width() / 2,
                                bar.get_height() * 1.02,
                                f"{pct:+.0f}%",
                                ha="center", va="bottom",
                                fontsize=6.5, color=col)

        ax.set_xticks(xs)
        ax.set_xticklabels(mnames, color=C["text_secondary"])
        ax.set_ylabel("Ricavi (€)")
        ax.set_title("Confronto Anno su Anno (YoY)", pad=14)
        ax.yaxis.set_major_formatter(
            mticker.FuncFormatter(lambda x, _: f"€{x:,.0f}"))
        ax.legend(framealpha=0, labelcolor=C["text_secondary"], fontsize=10)
        fig.tight_layout()
        self._hover_bars_v(ax, all_yoy_bars, all_yoy_lbls, all_yoy_vals)
        self._draw_done()
        self._render_stats(df.groupby("_y")[rev_col].sum())

    def _draw_mom(self, df, date_col, rev_col) -> None:
        df = df.copy()
        df["_p"] = df[date_col].dt.to_period("M")
        monthly  = df.groupby("_p")[rev_col].sum().sort_index()
        if len(monthly) < 2:
            self._empty("Servono almeno 2 mesi per il confronto MoM.")
            return

        pct = monthly.pct_change() * 100
        xs  = np.arange(len(monthly))
        bar_colors = [
            (C["accent_green"] if (pct.iloc[i] >= 0) else C["accent_warm"])
            if i > 0 else C["accent"]
            for i in range(len(monthly))
        ]

        fig, ax = self._make_fig(figsize=(12, 5))
        self._style(ax)
        ax.grid(axis="y"); ax.grid(axis="x", visible=False)

        bars = ax.bar(xs, monthly.values, color=bar_colors, width=0.6)
        for i, (bar, _) in enumerate(zip(bars, monthly.values)):
            if i > 0 and not np.isnan(pct.iloc[i]):
                col = C["accent_green"] if pct.iloc[i] >= 0 else C["accent_warm"]
                ax.text(bar.get_x() + bar.get_width() / 2,
                        bar.get_height() * 1.02,
                        f"{pct.iloc[i]:+.1f}%",
                        ha="center", va="bottom",
                        fontsize=7, color=col)

        labels = [str(p) for p in monthly.index]
        step   = max(1, len(labels) // 12)
        ax.set_xticks(xs[::step])
        ax.set_xticklabels(labels[::step], rotation=40, ha="right",
                           color=C["text_secondary"], fontsize=9)
        ax.set_ylabel("Ricavi (€)")
        ax.set_title("Confronto Mese su Mese (MoM)", pad=14)
        ax.yaxis.set_major_formatter(
            mticker.FuncFormatter(lambda x, _: f"€{x:,.0f}"))
        fig.tight_layout()
        self._hover_bars_v(ax,
                           list(bars),
                           labels,
                           monthly.values.tolist())
        self._draw_done()
        self._render_stats(monthly, pct)

    def _draw_yoy_line(self, df, date_col, rev_col) -> None:
        df = df.copy()
        df["_y"] = df[date_col].dt.year
        df["_m"] = df[date_col].dt.month
        pivot = df.groupby(["_y","_m"])[rev_col].sum().unstack(level=0)
        if pivot.empty:
            self._empty("Dati insufficienti per YoY.")
            return

        years  = sorted(pivot.columns)
        months = list(range(1, 13))
        mnames = ["Gen","Feb","Mar","Apr","Mag","Giu",
                  "Lug","Ago","Set","Ott","Nov","Dic"]
        xs = np.arange(len(months))

        fig, ax = self._make_fig(figsize=(12, 5))
        self._style(ax)
        ax.grid(axis="y"); ax.grid(axis="x", visible=False)

        year_vals_map: dict = {}
        for i, year in enumerate(years):
            vals = [float(pivot.loc[m, year])
                    if m in pivot.index and year in pivot.columns
                    else 0.0 for m in months]
            year_vals_map[str(year)] = vals
            col   = C["accent"] if i == len(years) - 1 else CHART_COLORS[i % len(CHART_COLORS)]
            alpha = 1.0 if i == len(years) - 1 else 0.55
            ax.plot(xs, vals, color=col, linewidth=2.2, alpha=alpha,
                    marker="o", markersize=4, label=str(year))

        ax.set_xticks(xs)
        ax.set_xticklabels(mnames, color=C["text_secondary"])
        ax.set_ylabel("Ricavi (€)")
        ax.set_title("Confronto Anno su Anno — Linee", pad=14)
        ax.yaxis.set_major_formatter(
            mticker.FuncFormatter(lambda x, _: f"€{x:,.0f}"))
        ax.legend(framealpha=0, labelcolor=C["text_secondary"], fontsize=10)
        fig.tight_layout()
        self._hover_multi_lines(ax, mnames, year_vals_map)
        self._draw_done()
        self._render_stats(df.groupby("_y")[rev_col].sum())

    def _draw_mom_line(self, df, date_col, rev_col) -> None:
        df = df.copy()
        df["_p"] = df[date_col].dt.to_period("M")
        monthly = df.groupby("_p")[rev_col].sum().sort_index()
        if len(monthly) < 2:
            self._empty("Servono almeno 2 mesi per il confronto MoM.")
            return

        xs = np.arange(len(monthly))
        pct = monthly.pct_change() * 100

        fig, ax = self._make_fig(figsize=(12, 5))
        self._style(ax)
        ax.grid(axis="y"); ax.grid(axis="x", visible=False)

        ax.plot(xs, monthly.values, color=C["accent"], linewidth=2.5,
                marker="o", markersize=5, label="Ricavi mensili")
        ax.fill_between(xs, monthly.values, alpha=0.12, color=C["accent"])

        labels = [str(p) for p in monthly.index]
        step   = max(1, len(labels) // 12)
        ax.set_xticks(xs[::step])
        ax.set_xticklabels(labels[::step], rotation=40, ha="right",
                           color=C["text_secondary"], fontsize=9)
        ax.set_ylabel("Ricavi (€)")
        ax.set_title("Andamento Mensile — Linea", pad=14)
        ax.yaxis.set_major_formatter(
            mticker.FuncFormatter(lambda x, _: f"€{x:,.0f}"))
        ax.legend(framealpha=0, labelcolor=C["text_secondary"], fontsize=10)
        fig.tight_layout()
        self._hover_line(ax, labels, monthly.values)
        self._draw_done()
        self._render_stats(monthly, pct)

    def _render_stats(self, series, pct_changes=None) -> None:
        bar = tk.Frame(self, bg=C["surface"])
        bar.pack(fill="x", padx=12, pady=(0, 8))
        self._stats_bar = bar

        def stat(label, val, color):
            cell = tk.Frame(bar, bg=C["surface"], padx=16, pady=8)
            cell.pack(side="left")
            tk.Label(cell, text=label, bg=C["surface"],
                     fg=C["text_secondary"],
                     font=(UI_FONT, 9)).pack(anchor="w")
            tk.Label(cell, text=val, bg=C["surface"], fg=color,
                     font=(MONO_FONT, 13, "bold")).pack(anchor="w")

        stat(f"Miglior periodo  {series.idxmax()}",
             f"€{series.max():,.0f}", C["accent_green"])
        tk.Frame(bar, bg=C["border"], width=1).pack(
            side="left", fill="y", pady=6)
        stat(f"Peggior periodo  {series.idxmin()}",
             f"€{series.min():,.0f}", C["accent_warm"])
        if pct_changes is not None:
            valid = pct_changes.dropna()
            if len(valid):
                avg = valid.mean()
                tk.Frame(bar, bg=C["border"], width=1).pack(
                    side="left", fill="y", pady=6)
                col = C["accent_green"] if avg >= 0 else C["accent_warm"]
                stat("Crescita media", f"{avg:+.1f}%", col)


# ═══════════════════════════════════════════════════════════════════════════════
# View 4 — Geo Map
# ═══════════════════════════════════════════════════════════════════════════════

class GeoMapView(BaseView):
    VIEW_TITLE    = "Mappa Geografica"
    VIEW_KEY      = "geo"
    DEFAULT_CHART = "scatter"

    def __init__(self, parent, app):
        super().__init__(parent, app)

    def refresh(self, df: pd.DataFrame, mapping: Dict) -> None:
        self._clear()
        geo_col = mapping.get("geo")
        if not geo_col:
            self._no_geo_state()
            return
        if df is None or df.empty:
            self._empty("Nessun dato da visualizzare.")
            return
        self._draw(df, mapping, geo_col)

    def _no_geo_state(self) -> None:
        """Proper empty state when no geographic column exists."""
        outer = tk.Frame(self._chart, bg=C["chart_bg"])
        outer.place(relx=0.5, rely=0.5, anchor="center")

        tk.Label(outer, text="🗺️", bg=C["chart_bg"],
                 font=(UI_FONT, 40)).pack(pady=(0, 10))
        tk.Label(outer, text="Nessuna colonna geografica",
                 bg=C["chart_bg"], fg=C["text_primary"],
                 font=(UI_FONT, 14, "bold")).pack()
        tk.Label(outer,
                 text="Aggiungi una colonna città o regione al tuo Excel.",
                 bg=C["chart_bg"], fg=C["text_secondary"],
                 font=(UI_FONT, 11)).pack(pady=(4, 12))
        detail = ("Nomi colonna riconosciuti:\n"
                  "  città · city · regione · region · provincia")
        tk.Label(outer, text=detail,
                 bg=C["surface"], fg=C["text_secondary"],
                 font=(UI_FONT, 10), justify="left",
                 padx=14, pady=10).pack()

    def _draw(self, df: pd.DataFrame, mapping: Dict, geo_col: str) -> None:
        rev_col  = mapping["revenue"]
        city_rev = df.groupby(geo_col)[rev_col].sum()

        matched: Dict[str, Tuple[float, float, float]] = {}
        unmatched: List[str] = []
        for city, rev in city_rev.items():
            key   = str(city).strip()
            found = next((k for k in ITALIAN_CITIES
                          if k.lower() == key.lower()), None)
            if found:
                matched[found] = (*ITALIAN_CITIES[found], float(rev))
            else:
                unmatched.append(key)

        fig, ax = self._make_fig(figsize=(10, 8))
        ax.set_facecolor("#0a0f1a")
        self._fig.patch.set_facecolor(C["chart_bg"])
        ax.tick_params(colors=C["text_secondary"], labelsize=8)
        for sp in ax.spines.values():
            sp.set_color(C["border"])
        ax.set_xlim(6.5, 18.6); ax.set_ylim(36.5, 47.2)
        ax.set_aspect("equal")
        ax.set_title("Distribuzione Geografica Ricavi — Italia", pad=14)
        ax.set_xlabel("Longitudine"); ax.set_ylabel("Latitudine")
        ax.grid(True, color=C["chart_grid"], linewidth=0.3,
                linestyle="--", alpha=0.5)

        if not matched:
            ax.text(12.5, 42,
                    "Nessuna città riconosciuta.\n"
                    "Controlla i nomi (es. Roma, Milano, Napoli).",
                    color=C["text_secondary"], fontsize=9,
                    ha="center", va="center",
                    bbox=dict(boxstyle="round,pad=0.4",
                              facecolor=C["surface"], alpha=0.7))
            self._draw_done()
            return

        max_r = max(v[2] for v in matched.values())
        min_r = min(v[2] for v in matched.values())
        rng   = max_r - min_r if max_r != min_r else 1

        for i, (city, (lat, lon, rev)) in enumerate(
                sorted(matched.items(), key=lambda x: x[1][2], reverse=True)):
            sz  = 60 + 700 * (rev - min_r) / rng
            col = CHART_COLORS[i % len(CHART_COLORS)]
            ax.scatter(lon, lat, s=sz, c=col, alpha=0.8,
                       edgecolors="white", linewidths=0.4, zorder=3)
            ax.annotate(
                f"{city}\n€{rev:,.0f}", (lon, lat),
                textcoords="offset points", xytext=(6, 4),
                fontsize=6.5, color=C["text_primary"],
                bbox=dict(boxstyle="round,pad=0.2",
                          facecolor=C["surface"], alpha=0.7))

        if unmatched:
            ax.set_xlabel(
                f"Longitudine  ·  Città non riconosciute ({len(unmatched)}): "
                + ", ".join(unmatched[:5])
                + ("…" if len(unmatched) > 5 else ""),
                fontsize=8)

        fig.tight_layout()
        self._draw_done()


# ═══════════════════════════════════════════════════════════════════════════════
# Main Application
# ═══════════════════════════════════════════════════════════════════════════════

class SalesAnalyzerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        _detect_fonts()     # must run after Tk() init

        self.title(f"Sales Analyzer v{APP_VERSION}")
        self.configure(bg=C["bg"])
        self.minsize(1060, 680)
        self.geometry("1300x820")
        try:
            self.iconbitmap(str(BASE_DIR / "icon.ico"))
        except Exception:
            pass

        self._cfg     = ConfigManager()
        self._mapper  = ColumnMapper()
        self._df_raw:       Optional[pd.DataFrame]             = None
        self._df_filtered:  Optional[pd.DataFrame]             = None
        self._mapping:      Optional[Dict[str, Optional[str]]] = None
        self._cur_view      = "trend"
        self._update_dismissed = False
        self._loading_widget:  Optional[tk.Widget]  = None
        self._manual_mapping   = False
        self._current_filename = ""
        self._chart_panel:     Optional[ChartTypePanel] = None

        # Per-view chart types (restored from config or default)
        saved_ct = self._cfg.get("chart_types", {})
        self._chart_types: Dict[str, str] = {
            k: saved_ct.get(k, v)
            for k, v in DEFAULT_CHART_TYPES.items()
        }

        self._build_layout()
        self._check_updates_bg()

    # ── Layout ────────────────────────────────────────────────────────────────

    def _build_layout(self) -> None:
        # ── Top bar ───────────────────────────────────────────────────────────
        self._topbar = TopBar(self, on_load=self._load_excel)
        self._topbar.pack(fill="x")

        # ── Body (sidebar + right panel) ──────────────────────────────────────
        body = tk.Frame(self, bg=C["bg"])
        body.pack(fill="both", expand=True)

        self._sidebar = Sidebar(body, on_select=self._switch_view)
        self._sidebar.pack(side="left", fill="y")

        self._right = tk.Frame(body, bg=C["bg"])
        self._right.pack(side="left", fill="both", expand=True)

        # KPI strip
        self._kpi = KPIStrip(self._right)
        self._kpi.pack(fill="x", padx=12, pady=(10, 0))

        # Filter bar
        self._fbar = FilterBar(self._right,
                               on_apply=self._apply_filters,
                               on_reset=self._reset_filters,
                               on_col_config=self._open_col_config,
                               on_chart_type=self._toggle_chart_panel)
        self._fbar.pack(fill="x", pady=(10, 0))

        # Chart wrapper (views live here)
        self._chart_wrap = tk.Frame(self._right, bg=C["bg"])
        self._chart_wrap.pack(fill="both", expand=True,
                              padx=0, pady=(0, 0))

        # Build views
        self._views: Dict[str, BaseView] = {
            "trend":      TrendView(self._chart_wrap, self),
            "top":        TopProductsView(self._chart_wrap, self),
            "comparison": PeriodComparisonView(self._chart_wrap, self),
            "geo":        GeoMapView(self._chart_wrap, self),
        }
        for v in self._views.values():
            v.place(relx=0, rely=0, relwidth=1, relheight=1)

        # Bottom bar: export button
        bbar = tk.Frame(self._right, bg=C["surface"], height=38)
        bbar.pack(fill="x")
        bbar.pack_propagate(False)
        tk.Frame(bbar, bg=C["border"], height=1).pack(side="top", fill="x")
        _Btn(bbar, "  ⬇  Esporta PNG  ",
             command=self._export_current,
             bg=C["surface"], fg=C["text_secondary"],
             font_args=(UI_FONT, 9), padx=14, pady=6).pack(
            side="right", padx=12)

        # Initial state: drop zone
        self._show_dropzone()
        self._switch_view("trend", animate=False)

    # ── Drop zone / loading states ────────────────────────────────────────────

    def _show_dropzone(self) -> None:
        dz = DropZone(self._chart_wrap, on_browse=self._load_excel)
        dz.place(relx=0, rely=0, relwidth=1, relheight=1)
        dz.lift()
        self._loading_widget = dz

    def _show_loading(self) -> None:
        if self._loading_widget:
            try:
                self._loading_widget.destroy()
            except Exception:
                pass
        lp = LoadingPlaceholder(self._chart_wrap)
        lp.place(relx=0, rely=0, relwidth=1, relheight=1)
        lp.lift()
        self._loading_widget = lp

    def _hide_loading(self) -> None:
        if self._loading_widget:
            try:
                self._loading_widget.destroy()
            except Exception:
                pass
            self._loading_widget = None

    # ── Excel loading ─────────────────────────────────────────────────────────

    def _load_excel(self) -> None:
        init = self._cfg.get("last_folder", str(Path.home()))
        path = filedialog.askopenfilename(
            title="Apri file Excel",
            initialdir=init,
            filetypes=[("Excel", "*.xlsx *.xls *.xlsm"), ("Tutti", "*.*")],
        )
        if not path:
            return
        self._cfg.set("last_folder", str(Path(path).parent))
        self._show_loading()

        def _load():
            try:
                df = pd.read_excel(path, engine="openpyxl")
                self.after(0, lambda: self._on_loaded(df, Path(path).name))
            except Exception as exc:
                log.error(f"Excel load error: {exc}")
                self.after(0, lambda: self._on_load_error(str(exc)))

        threading.Thread(target=_load, daemon=True).start()

    def _on_loaded(self, df: pd.DataFrame, filename: str) -> None:
        self._hide_loading()
        self._current_filename = filename

        # Try to parse date-like object columns
        for col in df.columns:
            if df[col].dtype == object:
                try:
                    df[col] = pd.to_datetime(df[col], dayfirst=True,
                                             infer_datetime_format=True)
                except Exception:
                    pass

        columns = list(df.columns)

        # Check for a previously saved manual mapping for this filename
        saved_mappings = self._cfg.get("column_mapping", {})
        saved = saved_mappings.get(filename)
        if saved and all(saved.get(r) is None or saved.get(r) in columns
                         for r in COLUMN_SYNONYMS):
            mapping = saved
            self._manual_mapping = True
        else:
            mapping = self._mapper.auto_map(columns)
            self._manual_mapping = False

        if not self._mapper.is_complete(mapping):
            dlg = ColumnMappingDialog(self, columns, mapping)
            self.wait_window(dlg)
            if dlg.result is None:
                self._show_dropzone()
                return
            mapping = dlg.result

        # Coerce types
        date_col = mapping["date"]
        try:
            df[date_col] = pd.to_datetime(df[date_col], dayfirst=True)
        except Exception as e:
            messagebox.showerror("Colonna Data",
                                 f"Impossibile interpretare la colonna data:\n{e}")
            self._show_dropzone()
            return

        for role in ("revenue", "quantity"):
            col = mapping.get(role)
            if col and col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        self._df_raw      = df
        self._mapping     = mapping
        self._df_filtered = df.copy()

        # Restore saved chart types to view instances
        for vk, ct in self._chart_types.items():
            if vk in self._views:
                self._views[vk]._chart_type = ct

        # Populate filter bar
        ch_col = mapping.get("channel")
        if ch_col and ch_col in df.columns:
            self._fbar.channel_dd.set_options(
                df[ch_col].dropna().unique().tolist())

        dmin = df[date_col].min().date()
        dmax = df[date_col].max().date()
        self._fbar.set_date(self._fbar.date_from, dmin)
        self._fbar.set_date(self._fbar.date_to,   dmax)

        self._fbar.set_buttons_enabled(True)
        self._fbar.set_col_badge(self._manual_mapping)

        self.title(f"Sales Analyzer v{APP_VERSION} — {filename}")
        self._topbar.set_filename(filename)

        self._update_kpis()
        self._switch_view(self._cur_view, animate=False)

    def _on_load_error(self, msg: str) -> None:
        self._hide_loading()
        self._show_dropzone()
        messagebox.showerror(
            "Errore Caricamento",
            f"Impossibile aprire il file.\n\n{msg}\n\n"
            "Verifica:\n"
            "  • File .xlsx valido e non aperto in Excel\n"
            "  • Almeno una colonna data, una prodotto, una ricavo")

    # ── Filters ───────────────────────────────────────────────────────────────

    def _apply_filters(self) -> None:
        if self._df_raw is None:
            messagebox.showinfo("Filtri", "Carica prima un file Excel.")
            return
        df       = self._df_raw.copy()
        date_col = self._mapping["date"]
        ch_col   = self._mapping.get("channel")

        d_from = self._fbar.get_date(self._fbar.date_from)
        d_to   = self._fbar.get_date(self._fbar.date_to)
        if d_from:
            df = df[df[date_col].dt.date >= d_from]
        if d_to:
            df = df[df[date_col].dt.date <= d_to]

        if ch_col and ch_col in df.columns:
            sel = self._fbar.channel_dd.get_selected()
            if sel:
                df = df[df[ch_col].isin(sel)]

        if df.empty:
            messagebox.showwarning("Filtri",
                                   "Nessun dato corrisponde ai filtri.")
            return

        self._df_filtered = df
        self._update_kpis()
        self._views[self._cur_view].refresh(df, self._mapping)

    def _reset_filters(self) -> None:
        if self._df_raw is None:
            return
        self._df_filtered = self._df_raw.copy()
        date_col = self._mapping["date"]
        dmin     = self._df_raw[date_col].min().date()
        dmax     = self._df_raw[date_col].max().date()
        self._fbar.set_date(self._fbar.date_from, dmin)
        self._fbar.set_date(self._fbar.date_to,   dmax)
        ch_col = self._mapping.get("channel")
        if ch_col:
            self._fbar.channel_dd.set_options(
                self._df_raw[ch_col].dropna().unique().tolist())
        self._update_kpis()
        self._views[self._cur_view].refresh(self._df_filtered, self._mapping)

    # ── KPI computation ───────────────────────────────────────────────────────

    def _compute_kpis(self):
        df  = self._df_filtered
        raw = self._df_raw
        m   = self._mapping
        if df is None or df.empty or m is None:
            return None

        dc = m["date"]; rc = m["revenue"]
        qc = m.get("quantity"); cc = m.get("channel")

        total_rev = float(df[rc].sum())
        total_qty = float(df[qc].sum()) if qc else 0.0
        n_orders  = len(df)
        avg_order = total_rev / n_orders if n_orders > 0 else 0.0

        top_ch     = "N/A"
        top_ch_pct = 0.0
        if cc and cc in df.columns:
            ch_rev = df.groupby(cc)[rc].sum()
            if not ch_rev.empty:
                top_ch     = str(ch_rev.idxmax())
                top_ch_pct = (float(ch_rev.max()) / total_rev * 100
                              if total_rev > 0 else 0.0)

        # Delta: compare vs same-length period immediately before
        dmin  = df[dc].min(); dmax = df[dc].max()
        span  = dmax - dmin
        p_end = dmin; p_start = p_end - span
        prev  = raw[(raw[dc] >= p_start) & (raw[dc] < p_end)]
        pr    = float(prev[rc].sum()) if not prev.empty else 0.0
        pq    = float(prev[qc].sum()) if qc and not prev.empty else 0.0

        rev_d = (total_rev - pr) / pr * 100 if pr > 0 else None
        qty_d = (total_qty - pq) / pq * 100 if pq > 0 else None

        return [
            ("₿",  "Ricavo Totale", total_rev, rev_d,     "currency"),
            ("📦", "Unità Vendute", total_qty, qty_d,     "count"),
            ("⌀",  "Ordine Medio",  avg_order, None,      "currency"),
            ("📡", "Canale Top",    top_ch,    top_ch_pct,"channel"),
        ]

    def _update_kpis(self) -> None:
        kpis = self._compute_kpis()
        if kpis:
            self._kpi.update(kpis)

    # ── View switching ────────────────────────────────────────────────────────

    def _switch_view(self, key: str, animate: bool = True) -> None:
        if key == self._cur_view and animate:
            return
        # Close floating chart panel if open
        if self._chart_panel and self._chart_panel.winfo_exists():
            self._chart_panel._close()

        def _do_switch() -> None:
            self._cur_view = key
            self._views[key].lift()
            self._sidebar.set_active(key)
            if self._df_filtered is not None and self._mapping:
                self._views[key].refresh(self._df_filtered, self._mapping)

        if not animate:
            _do_switch()
            return

        _fade_overlay(self._chart_wrap, _do_switch, duration_ms=120)

    # ── Feature 1: Column config ──────────────────────────────────────────────

    def _open_col_config(self) -> None:
        if self._df_raw is None:
            return
        dlg = ColumnConfigDialog(self, list(self._df_raw.columns),
                                  self._mapping, self._mapper)
        self.wait_window(dlg)
        if dlg.result:
            self._mapping      = dlg.result
            self._manual_mapping = True
            # Save mapping keyed by filename
            saved = self._cfg.get("column_mapping", {}) or {}
            saved[self._current_filename] = dlg.result
            self._cfg.set("column_mapping", saved)
            self._fbar.set_col_badge(True)
            # Re-coerce types with new mapping
            date_col = self._mapping.get("date")
            if date_col:
                try:
                    self._df_raw[date_col] = pd.to_datetime(
                        self._df_raw[date_col], dayfirst=True)
                except Exception:
                    pass
            for role in ("revenue", "quantity"):
                col = self._mapping.get(role)
                if col and col in self._df_raw.columns:
                    self._df_raw[col] = pd.to_numeric(
                        self._df_raw[col], errors="coerce").fillna(0)
            self._df_filtered = self._df_raw.copy()
            self._update_kpis()
            self._views[self._cur_view].refresh(self._df_filtered, self._mapping)
        elif dlg.reset_to_auto:
            self._manual_mapping = False
            saved = self._cfg.get("column_mapping", {}) or {}
            saved.pop(self._current_filename, None)
            self._cfg.set("column_mapping", saved)
            self._fbar.set_col_badge(False)
            # Re-run auto mapping
            self._mapping = self._mapper.auto_map(list(self._df_raw.columns))
            self._df_filtered = self._df_raw.copy()
            self._update_kpis()
            self._views[self._cur_view].refresh(self._df_filtered, self._mapping)

    # ── Feature 2: Chart type panel ───────────────────────────────────────────

    def _toggle_chart_panel(self, anchor: tk.Widget) -> None:
        if self._chart_panel and self._chart_panel.winfo_exists():
            self._chart_panel._close()
            return
        ct = self._chart_types.get(self._cur_view,
                                    DEFAULT_CHART_TYPES.get(self._cur_view, "line"))
        self._chart_panel = ChartTypePanel(
            self, anchor,
            view_key=self._cur_view,
            current_type=ct,
            on_select=self._set_chart_type,
            on_close=lambda: setattr(self, "_chart_panel", None),
        )

    def _set_chart_type(self, view_key: str, ct: str) -> None:
        self._chart_types[view_key] = ct
        saved = self._cfg.get("chart_types", {}) or {}
        saved[view_key] = ct
        self._cfg.set("chart_types", saved)
        view = self._views.get(view_key)
        if view:
            view.set_chart_type(ct)
        label = next((l for k, _, l in CHART_TYPES if k == ct), ct)
        self._show_toast(f"Grafico: {label}")

    def _show_toast(self, message: str) -> None:
        Toast(self, message, duration_ms=2000)

    # ── Export ────────────────────────────────────────────────────────────────

    def _export_current(self) -> None:
        self._views[self._cur_view].export_png()

    # ── Auto-updater ──────────────────────────────────────────────────────────

    def _check_updates_bg(self) -> None:
        try:
            from updater import check_for_update
            check_for_update(
                on_update_available=self._on_update_available,
                on_error=None,
            )
        except Exception as e:
            log.debug(f"Updater unavailable: {e}")

    def _on_update_available(self, version: str, url: str) -> None:
        if self._update_dismissed:
            return
        self.after(0, lambda: self._topbar.show_update_btn(
            version, lambda: self._start_download(version, url)))

    def _start_download(self, version: str, url: str) -> None:
        if not url:
            messagebox.showerror("Download",
                                 "URL non disponibile per questa release.")
            return
        dlg = DownloadDialog(self)

        def on_progress(d, t): dlg.update_progress(d, t)
        def on_error(m): dlg.show_error(m,
                           on_retry=lambda: self._start_download(version, url))
        def on_complete(path):
            dlg.show_complete()
            self.after(3000, lambda: _apply(path))

        def _apply(path):
            try:
                from updater import apply_update
                apply_update(path)
            except Exception as e:
                messagebox.showerror("Errore", str(e))

        try:
            from updater import download_update
            download_update(url, on_progress, on_complete, on_error)
        except Exception as e:
            on_error(str(e))


# ═══════════════════════════════════════════════════════════════════════════════
# Entry point
# ═══════════════════════════════════════════════════════════════════════════════

def main() -> None:
    app = SalesAnalyzerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
