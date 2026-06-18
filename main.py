#!/usr/bin/env python3
"""
Sales Analyzer — Desktop Sales Dashboard
Run: python main.py
"""

import os
import sys
import json
import logging
import threading
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
import matplotlib.patches as mpatches

# ── Optional calendar widget ──────────────────────────────────────────────────
try:
    from tkcalendar import DateEntry
    HAS_CALENDAR = True
except ImportError:
    HAS_CALENDAR = False

# ── Logging ───────────────────────────────────────────────────────────────────
BASE_DIR = Path(os.path.dirname(os.path.abspath(__file__)))
LOG_FILE = BASE_DIR / "app.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger(__name__)

# ── App version ───────────────────────────────────────────────────────────────
try:
    APP_VERSION = (BASE_DIR / "VERSION").read_text().strip()
except Exception:
    APP_VERSION = "1.0.0"

# ── Colour palette ─────────────────────────────────────────────────────────────
C = {
    "bg":         "#1a1a2e",
    "sidebar":    "#16213e",
    "card":       "#0f3460",
    "accent":     "#e94560",
    "accent2":    "#533483",
    "text":       "#eaeaea",
    "dim":        "#8888aa",
    "success":    "#4ade80",
    "warning":    "#fbbf24",
    "error":      "#f87171",
    "border":     "#2a2a4e",
    "hover":      "#2a2a5e",
    "chart_bg":   "#16213e",
    "chart_grid": "#2a2a4e",
}

CHART_COLORS = [
    "#e94560", "#4ecdc4", "#45b7d1", "#96ceb4", "#fbbf24",
    "#dda0dd", "#98d8c8", "#bb8fce", "#85c1e9", "#f9ca24",
]

# ── Column name synonyms ───────────────────────────────────────────────────────
COLUMN_SYNONYMS: Dict[str, List[str]] = {
    "date": [
        "data", "date", "periodo", "period", "mese", "month",
        "data_vendita", "data vendita", "data_ordine", "order_date",
        "giorno", "day", "timestamp", "anno", "year",
    ],
    "product": [
        "prodotto", "product", "categoria", "category", "articolo",
        "item", "nome_prodotto", "nome prodotto", "sku", "descrizione",
        "description", "nome",
    ],
    "revenue": [
        "ricavo", "revenue", "fatturato", "vendite", "importo",
        "amount", "totale", "total", "prezzo", "valore", "value",
        "incasso", "guadagno", "entrata", "sales", "price",
        "ricavi", "revenues", "importo_totale",
    ],
    "quantity": [
        "quantita", "quantity", "qty", "pezzi", "unita", "units",
        "quantità", "num", "numero", "count", "sold", "venduti",
        "q.ta", "qta",
    ],
    "channel": [
        "canale", "channel", "canale_vendita", "sales_channel",
        "tipo_vendita", "tipo", "type", "mezzo", "source",
        "origine", "modalita", "modalità",
    ],
    "geo": [
        "citta", "city", "città", "regione", "region",
        "provincia", "province", "luogo", "location",
        "area", "zona", "zone", "territorio",
    ],
}

ITALIAN_CITIES: Dict[str, Tuple[float, float]] = {
    "Roma": (41.9028, 12.4964), "Milano": (45.4654, 9.1859),
    "Napoli": (40.8518, 14.2681), "Torino": (45.0703, 7.6869),
    "Palermo": (38.1157, 13.3615), "Genova": (44.4056, 8.9463),
    "Bologna": (44.4949, 11.3426), "Firenze": (43.7696, 11.2558),
    "Bari": (41.1171, 16.8719), "Catania": (37.5079, 15.0830),
    "Venezia": (45.4408, 12.3155), "Verona": (45.4384, 10.9916),
    "Messina": (38.1938, 15.5540), "Padova": (45.4064, 11.8768),
    "Trieste": (45.6495, 13.7768), "Brescia": (45.5416, 10.2118),
    "Taranto": (40.4640, 17.2470), "Modena": (44.6471, 10.9252),
    "Reggio Calabria": (38.1113, 15.6474), "Reggio Emilia": (44.6989, 10.6297),
    "Perugia": (43.1121, 12.3888), "Ravenna": (44.4184, 12.2035),
    "Livorno": (43.5485, 10.3106), "Cagliari": (39.2238, 9.1217),
    "Foggia": (41.4600, 15.5440), "Rimini": (44.0678, 12.5695),
    "Salerno": (40.6824, 14.7681), "Ferrara": (44.8381, 11.6198),
    "Sassari": (40.7259, 8.5556), "Bergamo": (45.6983, 9.6773),
}


# ═══════════════════════════════════════════════════════════════════════════════
# Helpers
# ═══════════════════════════════════════════════════════════════════════════════

class ConfigManager:
    def __init__(self):
        self._path = BASE_DIR / "config.json"
        self._data: Dict[str, Any] = {}
        self._load()

    def _load(self):
        try:
            if self._path.exists():
                self._data = json.loads(self._path.read_text())
        except Exception as e:
            log.warning(f"Config load error: {e}")

    def save(self):
        try:
            self._path.write_text(json.dumps(self._data, indent=2))
        except Exception as e:
            log.warning(f"Config save error: {e}")

    def get(self, key: str, default=None):
        return self._data.get(key, default)

    def set(self, key: str, value):
        self._data[key] = value
        self.save()


class ColumnMapper:
    """Auto-detect which Excel column maps to each semantic role."""

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
        """Date, product, and revenue are required."""
        required = ("date", "product", "revenue")
        return all(mapping.get(r) for r in required)


# ═══════════════════════════════════════════════════════════════════════════════
# Custom Widgets
# ═══════════════════════════════════════════════════════════════════════════════

class MultiSelectDropdown(tk.Frame):
    """A button that opens a popup with checkboxes for multi-selection."""

    def __init__(self, parent, label="Canale", **kwargs):
        bg = kwargs.pop("bg", C["bg"])
        super().__init__(parent, bg=bg, **kwargs)
        self.options: List[str] = []
        self.vars: Dict[str, tk.BooleanVar] = {}
        self._popup: Optional[tk.Toplevel] = None
        self._label = label

        self._btn = tk.Button(
            self, text=f"{label}: Tutti ▼",
            bg=C["card"], fg=C["text"],
            relief="flat", bd=0,
            padx=10, pady=4,
            font=("Segoe UI", 9),
            cursor="hand2",
            activebackground=C["hover"], activeforeground=C["text"],
            command=self._toggle,
        )
        self._btn.pack()

    def set_options(self, options: List[str]):
        self.options = sorted(options)
        self.vars = {opt: tk.BooleanVar(value=True) for opt in self.options}
        self._update_label()

    def get_selected(self) -> List[str]:
        return [k for k, v in self.vars.items() if v.get()]

    def _update_label(self):
        sel = self.get_selected()
        if not self.options:
            self._btn.config(text=f"{self._label}: Tutti ▼")
        elif len(sel) == len(self.options):
            self._btn.config(text=f"{self._label}: Tutti ▼")
        elif len(sel) == 0:
            self._btn.config(text=f"{self._label}: Nessuno ▼")
        else:
            self._btn.config(text=f"{self._label}: {len(sel)} sel. ▼")

    def _toggle(self):
        if self._popup and self._popup.winfo_exists():
            self._popup.destroy()
            self._popup = None
            return
        self._open_popup()

    def _open_popup(self):
        self._popup = tk.Toplevel(self, bg=C["card"])
        self._popup.wm_overrideredirect(True)
        self._popup.wm_attributes("-topmost", True)

        x = self._btn.winfo_rootx()
        y = self._btn.winfo_rooty() + self._btn.winfo_height() + 2
        self._popup.geometry(f"+{x}+{y}")

        # "All" toggle button
        def toggle_all():
            all_on = all(v.get() for v in self.vars.values())
            for v in self.vars.values():
                v.set(not all_on)
            self._update_label()

        tk.Button(
            self._popup, text="Seleziona Tutti / Nessuno",
            bg=C["accent"], fg=C["text"], relief="flat",
            command=toggle_all, padx=6, pady=3,
            font=("Segoe UI", 8),
        ).pack(fill="x", padx=4, pady=(4, 2))

        for opt in self.options:
            var = self.vars[opt]
            cb = tk.Checkbutton(
                self._popup, text=opt, variable=var,
                bg=C["card"], fg=C["text"],
                selectcolor=C["accent2"],
                activebackground=C["card"],
                activeforeground=C["text"],
                font=("Segoe UI", 9),
                command=self._update_label,
            )
            cb.pack(anchor="w", padx=8, pady=1)

        # Close on outside click
        self._popup.bind("<FocusOut>", lambda e: self._close_popup())
        self._popup.focus_set()

    def _close_popup(self):
        if self._popup and self._popup.winfo_exists():
            self._popup.destroy()
            self._popup = None


class LoadingOverlay(tk.Frame):
    """A semi-transparent overlay with a progress indicator."""

    def __init__(self, parent):
        super().__init__(parent, bg=C["bg"])
        self._angle = 0
        self._canvas = tk.Canvas(self, width=60, height=60, bg=C["bg"], highlightthickness=0)
        self._canvas.pack(pady=(60, 10))
        tk.Label(self, text="Caricamento in corso…", bg=C["bg"],
                 fg=C["dim"], font=("Segoe UI", 11)).pack()
        self._animate()

    def _animate(self):
        self._canvas.delete("all")
        cx, cy, r = 30, 30, 20
        self._canvas.create_arc(
            cx - r, cy - r, cx + r, cy + r,
            start=self._angle, extent=270,
            outline=C["accent"], width=4, style="arc",
        )
        self._angle = (self._angle + 12) % 360
        self._job = self.after(40, self._animate)

    def destroy(self):
        try:
            self.after_cancel(self._job)
        except Exception:
            pass
        super().destroy()


class UpdateBanner(tk.Frame):
    """A dismissible yellow banner shown when an update is available."""

    def __init__(self, parent, version: str, on_download, **kwargs):
        super().__init__(parent, bg=C["warning"], **kwargs)
        msg = f"  Nuova versione disponibile: v{version} — "
        tk.Label(self, text=msg, bg=C["warning"], fg="#1a1a2e",
                 font=("Segoe UI", 9, "bold")).pack(side="left", pady=4)
        tk.Button(
            self, text="Scarica e aggiorna",
            bg=C["accent"], fg=C["text"],
            relief="flat", padx=8, pady=2,
            font=("Segoe UI", 9, "bold"),
            cursor="hand2",
            command=on_download,
        ).pack(side="left", pady=4)
        tk.Button(
            self, text="✕", bg=C["warning"], fg="#1a1a2e",
            relief="flat", padx=6, pady=2,
            font=("Segoe UI", 10, "bold"),
            cursor="hand2",
            command=self.destroy,
        ).pack(side="right", pady=4, padx=4)


# ═══════════════════════════════════════════════════════════════════════════════
# Column Mapping Dialog
# ═══════════════════════════════════════════════════════════════════════════════

class ColumnMappingDialog(tk.Toplevel):
    ROLES = {
        "date":     ("Data / Periodo *", True),
        "product":  ("Prodotto / Categoria *", True),
        "revenue":  ("Ricavo / Importo *", True),
        "quantity": ("Quantità Venduta", False),
        "channel":  ("Canale di Vendita", False),
        "geo":      ("Città / Regione", False),
    }

    def __init__(self, parent, columns: List[str], mapping: Dict[str, Optional[str]]):
        super().__init__(parent)
        self.title("Mappa Colonne Excel")
        self.configure(bg=C["bg"])
        self.resizable(False, False)
        self.grab_set()

        self.result: Optional[Dict[str, Optional[str]]] = None
        self._combos: Dict[str, ttk.Combobox] = {}

        choices = ["— Non presente —"] + list(columns)

        tk.Label(self, text="Associa le colonne del tuo Excel ai campi richiesti:",
                 bg=C["bg"], fg=C["text"], font=("Segoe UI", 10)).pack(pady=(16, 8), padx=20)

        form = tk.Frame(self, bg=C["bg"])
        form.pack(padx=20, pady=4)

        for row_i, (role, (label, required)) in enumerate(self.ROLES.items()):
            lbl_text = label if required else label + " (opzionale)"
            tk.Label(form, text=lbl_text, bg=C["bg"], fg=C["text"],
                     font=("Segoe UI", 9), anchor="w", width=28).grid(
                row=row_i, column=0, sticky="w", pady=4, padx=(0, 12))

            style = ttk.Style()
            style.configure("Dark.TCombobox",
                             fieldbackground=C["card"],
                             background=C["card"],
                             foreground=C["text"],
                             selectbackground=C["accent"])
            cb = ttk.Combobox(form, values=choices, state="readonly",
                               width=28, style="Dark.TCombobox")
            current = mapping.get(role)
            cb.set(current if current else "— Non presente —")
            cb.grid(row=row_i, column=1, pady=4)
            self._combos[role] = cb

        btn_frame = tk.Frame(self, bg=C["bg"])
        btn_frame.pack(pady=16)
        tk.Button(btn_frame, text="Annulla", bg=C["card"], fg=C["text"],
                  relief="flat", padx=16, pady=6,
                  command=self.destroy).pack(side="left", padx=8)
        tk.Button(btn_frame, text="Conferma", bg=C["accent"], fg=C["text"],
                  relief="flat", padx=16, pady=6,
                  command=self._confirm).pack(side="left", padx=8)

        self.update_idletasks()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_rootx(), parent.winfo_rooty()
        w, h = self.winfo_reqwidth(), self.winfo_reqheight()
        self.geometry(f"+{px + (pw - w) // 2}+{py + (ph - h) // 2}")

    def _confirm(self):
        result = {}
        for role, cb in self._combos.items():
            val = cb.get()
            result[role] = None if val.startswith("—") else val
        required = ("date", "product", "revenue")
        missing = [r for r in required if not result.get(r)]
        if missing:
            messagebox.showerror(
                "Campi Mancanti",
                f"I seguenti campi obbligatori non sono stati associati:\n"
                + "\n".join(f"  • {self.ROLES[r][0]}" for r in missing),
                parent=self,
            )
            return
        self.result = result
        self.destroy()


# ═══════════════════════════════════════════════════════════════════════════════
# Chart Views — base
# ═══════════════════════════════════════════════════════════════════════════════

class BaseView(tk.Frame):
    """Base class for all chart views."""

    VIEW_TITLE = "Vista"

    def __init__(self, parent, app: "SalesAnalyzerApp"):
        super().__init__(parent, bg=C["bg"])
        self.app = app
        self._current_fig: Optional[Figure] = None
        self._canvas_widget: Optional[FigureCanvasTkAgg] = None

        # Top bar: controls + export button
        self._controls_bar = tk.Frame(self, bg=C["bg"])
        self._controls_bar.pack(fill="x", padx=8, pady=(8, 0))

        self._export_btn = tk.Button(
            self._controls_bar, text="⬇ Esporta PNG",
            bg=C["card"], fg=C["text"],
            relief="flat", padx=10, pady=4,
            font=("Segoe UI", 9),
            cursor="hand2",
            activebackground=C["hover"], activeforeground=C["text"],
            command=self._export_png,
        )
        self._export_btn.pack(side="right", padx=4)

        # Chart container
        self._chart_frame = tk.Frame(self, bg=C["chart_bg"])
        self._chart_frame.pack(fill="both", expand=True, padx=8, pady=8)

    def _make_figure(self, nrows=1, ncols=1, **kwargs) -> Tuple[Figure, Any]:
        if self._canvas_widget:
            self._canvas_widget.get_tk_widget().destroy()
            self._canvas_widget = None

        fig = Figure(facecolor=C["chart_bg"], **kwargs)
        self._current_fig = fig
        canvas = FigureCanvasTkAgg(fig, master=self._chart_frame)
        canvas.get_tk_widget().pack(fill="both", expand=True)
        self._canvas_widget = canvas
        ax = fig.add_subplot(nrows, ncols, 1) if nrows * ncols == 1 else None
        return fig, ax

    def _style_ax(self, ax):
        ax.set_facecolor(C["chart_bg"])
        ax.tick_params(colors=C["dim"], labelsize=8)
        ax.spines["bottom"].set_color(C["border"])
        ax.spines["left"].set_color(C["border"])
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.grid(True, color=C["chart_grid"], linewidth=0.5, alpha=0.7)
        ax.title.set_color(C["text"])
        ax.xaxis.label.set_color(C["dim"])
        ax.yaxis.label.set_color(C["dim"])

    def _show_empty(self, message: str = "Nessun dato disponibile"):
        if self._canvas_widget:
            self._canvas_widget.get_tk_widget().destroy()
            self._canvas_widget = None
        lbl = tk.Label(
            self._chart_frame, text=message,
            bg=C["chart_bg"], fg=C["dim"],
            font=("Segoe UI", 12), wraplength=400, justify="center",
        )
        lbl.pack(expand=True)

    def _export_png(self):
        if not self._current_fig:
            messagebox.showinfo("Export", "Nessun grafico da esportare.", parent=self)
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png")],
            initialfile=f"{self.VIEW_TITLE.replace(' ', '_')}.png",
        )
        if path:
            self._current_fig.savefig(path, dpi=150, bbox_inches="tight",
                                      facecolor=C["chart_bg"])
            messagebox.showinfo("Export", f"Grafico salvato:\n{path}", parent=self)

    def refresh(self, df: pd.DataFrame, mapping: Dict[str, Optional[str]]):
        """Override in each view to render content."""
        raise NotImplementedError

    def _clear_chart_frame(self):
        for widget in self._chart_frame.winfo_children():
            widget.destroy()
        self._canvas_widget = None
        self._current_fig = None


# ═══════════════════════════════════════════════════════════════════════════════
# View 1 — Trend
# ═══════════════════════════════════════════════════════════════════════════════

class TrendView(BaseView):
    VIEW_TITLE = "Trend Ricavi"

    def __init__(self, parent, app):
        super().__init__(parent, app)
        self._granularity = tk.StringVar(value="Mese")
        self._show_qty = tk.BooleanVar(value=False)

        # Controls
        btn_frame = tk.Frame(self._controls_bar, bg=C["bg"])
        btn_frame.pack(side="left")

        tk.Label(btn_frame, text="Granularità:", bg=C["bg"], fg=C["dim"],
                 font=("Segoe UI", 9)).pack(side="left", padx=(0, 6))

        self._gran_btns: Dict[str, tk.Button] = {}
        for gran in ("Giorno", "Settimana", "Mese", "Trimestre"):
            btn = tk.Button(
                btn_frame, text=gran,
                bg=C["card"], fg=C["text"],
                relief="flat", padx=10, pady=4,
                font=("Segoe UI", 9),
                cursor="hand2",
                command=lambda g=gran: self._set_granularity(g),
            )
            btn.pack(side="left", padx=2)
            self._gran_btns[gran] = btn

        tk.Checkbutton(
            self._controls_bar, text="Mostra Quantità",
            variable=self._show_qty, bg=C["bg"], fg=C["text"],
            selectcolor=C["accent2"],
            activebackground=C["bg"], activeforeground=C["text"],
            font=("Segoe UI", 9),
            command=self._on_toggle_qty,
        ).pack(side="left", padx=12)

        self._df: Optional[pd.DataFrame] = None
        self._mapping: Optional[Dict] = None
        self._highlight_btn(self._granularity.get())

    def _highlight_btn(self, selected: str):
        for g, btn in self._gran_btns.items():
            btn.config(bg=C["accent"] if g == selected else C["card"])

    def _set_granularity(self, gran: str):
        self._granularity.set(gran)
        self._highlight_btn(gran)
        if self._df is not None:
            self._draw(self._df, self._mapping)

    def _on_toggle_qty(self):
        if self._df is not None:
            self._draw(self._df, self._mapping)

    def refresh(self, df: pd.DataFrame, mapping: Dict):
        self._df = df
        self._mapping = mapping
        self._draw(df, mapping)

    def _draw(self, df: pd.DataFrame, mapping: Dict):
        self._clear_chart_frame()
        if df is None or df.empty:
            self._show_empty("Nessun dato da visualizzare.")
            return

        date_col = mapping["date"]
        rev_col = mapping["revenue"]
        qty_col = mapping.get("quantity")

        gran = self._granularity.get()
        freq_map = {
            "Giorno": "D", "Settimana": "W",
            "Mese": "ME", "Trimestre": "QE",
        }
        freq = freq_map[gran]

        # Resample
        ts = df.set_index(date_col)[[rev_col]].copy()
        if qty_col and self._show_qty.get():
            ts[qty_col] = df.set_index(date_col)[qty_col]

        # Handle potential resample compatibility
        try:
            grouped = ts.resample(freq).sum()
        except Exception:
            freq_fallback = {"ME": "M", "QE": "Q"}.get(freq, freq)
            grouped = ts.resample(freq_fallback).sum()

        if grouped.empty:
            self._show_empty("Nessun dato per il periodo selezionato.")
            return

        fig, ax = self._make_figure(figsize=(10, 5))
        self._style_ax(ax)

        ax.plot(grouped.index, grouped[rev_col],
                color=C["accent"], linewidth=2, marker="o", markersize=3,
                label="Ricavi")
        ax.fill_between(grouped.index, grouped[rev_col],
                        alpha=0.15, color=C["accent"])
        ax.set_ylabel("Ricavi (€)", color=C["dim"])
        ax.set_title(f"Andamento Ricavi — {gran}", color=C["text"], pad=12)

        # Quantity on secondary axis
        if qty_col and self._show_qty.get() and qty_col in grouped.columns:
            ax2 = ax.twinx()
            ax2.set_facecolor(C["chart_bg"])
            ax2.tick_params(colors=C["dim"], labelsize=8)
            ax2.spines["top"].set_visible(False)
            ax2.spines["left"].set_visible(False)
            ax2.spines["bottom"].set_color(C["border"])
            ax2.spines["right"].set_color(C["border"])
            ax2.plot(grouped.index, grouped[qty_col],
                     color=C["success"], linewidth=1.5, linestyle="--",
                     marker="s", markersize=3, label="Quantità")
            ax2.set_ylabel("Quantità", color=C["dim"])
            lines2, labels2 = ax2.get_legend_handles_labels()
            ax.legend(
                *[x + y for x, y in zip(
                    zip(*ax.get_legend_handles_labels()),
                    zip(*[lines2, labels2])
                )],
                loc="upper left",
                framealpha=0.2, labelcolor=C["text"],
            )
        else:
            ax.legend(loc="upper left", framealpha=0.2, labelcolor=C["text"])

        ax.yaxis.set_major_formatter(mticker.FuncFormatter(
            lambda x, _: f"€{x:,.0f}"))
        fig.tight_layout()
        self._canvas_widget.draw()


# ═══════════════════════════════════════════════════════════════════════════════
# View 2 — Top Products
# ═══════════════════════════════════════════════════════════════════════════════

class TopProductsView(BaseView):
    VIEW_TITLE = "Top Prodotti"

    def __init__(self, parent, app):
        super().__init__(parent, app)
        self._top_n = tk.StringVar(value="10")

        btn_frame = tk.Frame(self._controls_bar, bg=C["bg"])
        btn_frame.pack(side="left")
        tk.Label(btn_frame, text="Top N:", bg=C["bg"], fg=C["dim"],
                 font=("Segoe UI", 9)).pack(side="left", padx=(0, 6))

        self._n_btns: Dict[str, tk.Button] = {}
        for n in ("5", "10", "20", "Tutti"):
            btn = tk.Button(
                btn_frame, text=n,
                bg=C["card"], fg=C["text"],
                relief="flat", padx=10, pady=4,
                font=("Segoe UI", 9), cursor="hand2",
                command=lambda v=n: self._set_n(v),
            )
            btn.pack(side="left", padx=2)
            self._n_btns[n] = btn

        self._df: Optional[pd.DataFrame] = None
        self._mapping: Optional[Dict] = None
        self._highlight_btn("10")

    def _highlight_btn(self, selected: str):
        for n, btn in self._n_btns.items():
            btn.config(bg=C["accent"] if n == selected else C["card"])

    def _set_n(self, n: str):
        self._top_n.set(n)
        self._highlight_btn(n)
        if self._df is not None:
            self._draw(self._df, self._mapping)

    def refresh(self, df: pd.DataFrame, mapping: Dict):
        self._df = df
        self._mapping = mapping
        self._draw(df, mapping)

    def _draw(self, df: pd.DataFrame, mapping: Dict):
        self._clear_chart_frame()
        if df is None or df.empty:
            self._show_empty("Nessun dato da visualizzare.")
            return

        prod_col = mapping["product"]
        rev_col = mapping["revenue"]
        cat_col = mapping.get("product")  # use category if separate

        agg = df.groupby(prod_col)[rev_col].sum().sort_values(ascending=False)

        n_str = self._top_n.get()
        if n_str != "Tutti":
            agg = agg.head(int(n_str))

        if agg.empty:
            self._show_empty("Nessun prodotto trovato.")
            return

        n = len(agg)
        fig_h = max(4, min(n * 0.45 + 1, 12))
        fig, ax = self._make_figure(figsize=(10, fig_h))
        self._style_ax(ax)
        ax.grid(axis="x", color=C["chart_grid"], linewidth=0.5, alpha=0.7)
        ax.grid(axis="y", visible=False)

        colors = [CHART_COLORS[i % len(CHART_COLORS)] for i in range(n)]
        bars = ax.barh(range(n), agg.values, color=colors, height=0.65)

        ax.set_yticks(range(n))
        ax.set_yticklabels(
            [str(p)[:30] for p in reversed(list(agg.index))]
            if False else [str(p)[:30] for p in agg.index],
            fontsize=8, color=C["text"],
        )
        ax.invert_yaxis()

        # Value labels
        for bar, val in zip(bars, agg.values):
            ax.text(
                bar.get_width() + agg.max() * 0.01, bar.get_y() + bar.get_height() / 2,
                f"€{val:,.0f}", va="center", ha="left",
                color=C["dim"], fontsize=7,
            )

        ax.set_xlabel("Ricavo Totale (€)", color=C["dim"])
        ax.set_title(f"Top {n_str} Prodotti per Ricavo", color=C["text"], pad=12)
        ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"€{x:,.0f}"))
        fig.tight_layout()
        self._canvas_widget.draw()


# ═══════════════════════════════════════════════════════════════════════════════
# View 3 — Period Comparison
# ═══════════════════════════════════════════════════════════════════════════════

class PeriodComparisonView(BaseView):
    VIEW_TITLE = "Confronto Periodi"

    def __init__(self, parent, app):
        super().__init__(parent, app)
        self._mode = tk.StringVar(value="YoY")

        btn_frame = tk.Frame(self._controls_bar, bg=C["bg"])
        btn_frame.pack(side="left")
        tk.Label(btn_frame, text="Confronto:", bg=C["bg"], fg=C["dim"],
                 font=("Segoe UI", 9)).pack(side="left", padx=(0, 6))

        self._mode_btns: Dict[str, tk.Button] = {}
        for mode in ("YoY", "MoM"):
            btn = tk.Button(
                btn_frame, text=mode,
                bg=C["card"], fg=C["text"],
                relief="flat", padx=12, pady=4,
                font=("Segoe UI", 9), cursor="hand2",
                command=lambda m=mode: self._set_mode(m),
            )
            btn.pack(side="left", padx=2)
            self._mode_btns[mode] = btn

        self._df: Optional[pd.DataFrame] = None
        self._mapping: Optional[Dict] = None
        self._stats_frame: Optional[tk.Frame] = None
        self._highlight_btn("YoY")

    def _highlight_btn(self, selected: str):
        for m, btn in self._mode_btns.items():
            btn.config(bg=C["accent"] if m == selected else C["card"])

    def _set_mode(self, mode: str):
        self._mode.set(mode)
        self._highlight_btn(mode)
        if self._df is not None:
            self._draw(self._df, self._mapping)

    def refresh(self, df: pd.DataFrame, mapping: Dict):
        self._df = df
        self._mapping = mapping
        self._draw(df, mapping)

    def _draw(self, df: pd.DataFrame, mapping: Dict):
        self._clear_chart_frame()
        if self._stats_frame:
            self._stats_frame.destroy()
            self._stats_frame = None

        if df is None or df.empty:
            self._show_empty("Nessun dato da visualizzare.")
            return

        date_col = mapping["date"]
        rev_col = mapping["revenue"]
        mode = self._mode.get()

        if mode == "YoY":
            self._draw_yoy(df, date_col, rev_col)
        else:
            self._draw_mom(df, date_col, rev_col)

    def _draw_yoy(self, df, date_col, rev_col):
        """Year-over-Year: each month Jan-Dec, bar per year."""
        df = df.copy()
        df["_year"] = df[date_col].dt.year
        df["_month"] = df[date_col].dt.month

        pivot = df.groupby(["_year", "_month"])[rev_col].sum().unstack(level=0)
        if pivot.empty:
            self._show_empty("Dati insufficienti per il confronto YoY.")
            return

        years = sorted(pivot.columns)
        months = list(range(1, 13))
        month_names = ["Gen", "Feb", "Mar", "Apr", "Mag", "Giu",
                       "Lug", "Ago", "Set", "Ott", "Nov", "Dic"]

        fig, ax = self._make_figure(figsize=(12, 5))
        self._style_ax(ax)
        ax.grid(axis="y", color=C["chart_grid"], linewidth=0.5, alpha=0.7)
        ax.grid(axis="x", visible=False)

        n_years = len(years)
        bar_w = 0.7 / n_years
        xs = np.arange(len(months))

        for i, year in enumerate(years):
            vals = [pivot.loc[m, year] if m in pivot.index and year in pivot.columns
                    else 0 for m in months]
            offset = (i - (n_years - 1) / 2) * bar_w
            bars = ax.bar(xs + offset, vals, bar_w * 0.9,
                          color=CHART_COLORS[i % len(CHART_COLORS)], label=str(year))

            # % change labels (compare to previous year same month)
            if i > 0:
                prev_year = years[i - 1]
                for j, (bar, m) in enumerate(zip(bars, months)):
                    curr = vals[j]
                    prev = pivot.loc[m, prev_year] if m in pivot.index and prev_year in pivot.columns else 0
                    if prev and curr:
                        pct = (curr - prev) / prev * 100
                        color = C["success"] if pct >= 0 else C["error"]
                        ax.text(
                            bar.get_x() + bar.get_width() / 2,
                            bar.get_height() + ax.get_ylim()[1] * 0.01,
                            f"{pct:+.0f}%", ha="center", va="bottom",
                            fontsize=6, color=color,
                        )

        ax.set_xticks(xs)
        ax.set_xticklabels(month_names, color=C["dim"])
        ax.set_ylabel("Ricavi (€)", color=C["dim"])
        ax.set_title("Confronto Anno su Anno (YoY)", color=C["text"], pad=12)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"€{x:,.0f}"))
        ax.legend(framealpha=0.2, labelcolor=C["text"])

        fig.tight_layout()
        self._canvas_widget.draw()
        self._render_stats_yoy(df, rev_col, "_year")

    def _draw_mom(self, df, date_col, rev_col):
        """Month-over-Month: each of last 12+ months, % change from prev."""
        df = df.copy()
        df["_period"] = df[date_col].dt.to_period("M")
        monthly = df.groupby("_period")[rev_col].sum().sort_index()

        if len(monthly) < 2:
            self._show_empty("Dati insufficienti per il confronto MoM (servono almeno 2 mesi).")
            return

        pct_changes = monthly.pct_change() * 100

        fig, ax = self._make_figure(figsize=(12, 5))
        self._style_ax(ax)
        ax.grid(axis="y", color=C["chart_grid"], linewidth=0.5, alpha=0.7)
        ax.grid(axis="x", visible=False)

        xs = np.arange(len(monthly))
        bar_colors = [C["success"] if (pct_changes.iloc[i] >= 0 if i > 0 else True)
                      else C["error"] for i in range(len(monthly))]
        bars = ax.bar(xs, monthly.values, color=bar_colors, width=0.6)

        for i, (bar, val) in enumerate(zip(bars, monthly.values)):
            if i > 0:
                pct = pct_changes.iloc[i]
                if not np.isnan(pct):
                    ax.text(
                        bar.get_x() + bar.get_width() / 2,
                        bar.get_height() + monthly.max() * 0.01,
                        f"{pct:+.1f}%", ha="center", va="bottom",
                        fontsize=7,
                        color=C["success"] if pct >= 0 else C["error"],
                    )

        labels = [str(p) for p in monthly.index]
        step = max(1, len(labels) // 12)
        ax.set_xticks(xs[::step])
        ax.set_xticklabels(labels[::step], rotation=45, ha="right",
                           color=C["dim"], fontsize=7)
        ax.set_ylabel("Ricavi (€)", color=C["dim"])
        ax.set_title("Confronto Mese su Mese (MoM)", color=C["text"], pad=12)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"€{x:,.0f}"))

        fig.tight_layout()
        self._canvas_widget.draw()
        self._render_stats_mom(monthly, pct_changes)

    def _render_stats_yoy(self, df, rev_col, year_col):
        yearly = df.groupby(year_col)[rev_col].sum()
        self._render_stats_panel(yearly)

    def _render_stats_mom(self, monthly, pct_changes):
        self._render_stats_panel(monthly, pct_changes)

    def _render_stats_panel(self, series, pct_changes=None):
        frame = tk.Frame(self, bg=C["sidebar"])
        frame.pack(fill="x", padx=8, pady=(0, 8))
        self._stats_frame = frame

        tk.Label(frame, text="Riepilogo Statistiche",
                 bg=C["sidebar"], fg=C["text"],
                 font=("Segoe UI", 9, "bold")).pack(side="left", padx=12, pady=6)

        best_label = str(series.idxmax())
        worst_label = str(series.idxmin())
        best_val = series.max()
        worst_val = series.min()

        stats = [
            (f"Miglior periodo: {best_label}", f"€{best_val:,.0f}", C["success"]),
            (f"Peggior periodo: {worst_label}", f"€{worst_val:,.0f}", C["error"]),
        ]
        if pct_changes is not None:
            valid = pct_changes.dropna()
            if len(valid):
                avg_growth = valid.mean()
                col = C["success"] if avg_growth >= 0 else C["error"]
                stats.append((f"Crescita media:", f"{avg_growth:+.1f}%", col))

        for label, val, color in stats:
            tk.Label(frame, text=f"  {label}", bg=C["sidebar"], fg=C["dim"],
                     font=("Segoe UI", 8)).pack(side="left", padx=4)
            tk.Label(frame, text=val, bg=C["sidebar"], fg=color,
                     font=("Segoe UI", 8, "bold")).pack(side="left", padx=(0, 12))


# ═══════════════════════════════════════════════════════════════════════════════
# View 4 — Geo Map
# ═══════════════════════════════════════════════════════════════════════════════

class GeoMapView(BaseView):
    VIEW_TITLE = "Mappa Geografica"

    def __init__(self, parent, app):
        super().__init__(parent, app)

    def refresh(self, df: pd.DataFrame, mapping: Dict):
        self._clear_chart_frame()
        geo_col = mapping.get("geo")

        if not geo_col:
            self._show_no_geo_column()
            return

        if df is None or df.empty:
            self._show_empty("Nessun dato da visualizzare.")
            return

        self._draw(df, mapping, geo_col)

    def _show_no_geo_column(self):
        outer = tk.Frame(self._chart_frame, bg=C["chart_bg"])
        outer.pack(expand=True)
        tk.Label(outer, text="📍", bg=C["chart_bg"], fg=C["dim"],
                 font=("Segoe UI", 32)).pack(pady=(40, 8))
        tk.Label(
            outer,
            text="Colonna geografica non trovata",
            bg=C["chart_bg"], fg=C["text"],
            font=("Segoe UI", 13, "bold"),
        ).pack()
        tk.Label(
            outer,
            text=(
                "Aggiungi al tuo Excel una colonna con nomi di città o regioni.\n"
                "Nomi colonna supportati: città, city, regione, region, provincia, province, luogo\n\n"
                "Esempio:\n"
                "  Città        → Roma, Milano, Napoli …\n"
                "  Regione      → Lazio, Lombardia, Campania …"
            ),
            bg=C["chart_bg"], fg=C["dim"],
            font=("Segoe UI", 10),
            justify="left",
        ).pack(pady=12, padx=40)

    def _draw(self, df: pd.DataFrame, mapping: Dict, geo_col: str):
        rev_col = mapping["revenue"]
        channel_col = mapping.get("channel")

        # Normalize city names for lookup
        city_rev = df.groupby(geo_col)[rev_col].sum()

        matched = {}
        unmatched = []
        for city, rev in city_rev.items():
            key = str(city).strip()
            found = None
            for k in ITALIAN_CITIES:
                if k.lower() == key.lower():
                    found = k
                    break
            if found:
                matched[found] = (ITALIAN_CITIES[found][0],
                                  ITALIAN_CITIES[found][1], rev)
            else:
                unmatched.append(key)

        fig, ax = self._make_figure(figsize=(10, 8))
        ax.set_facecolor("#0d1b2a")
        self._current_fig.patch.set_facecolor(C["chart_bg"])
        ax.tick_params(colors=C["dim"], labelsize=7)
        for spine in ax.spines.values():
            spine.set_color(C["border"])

        # Italy bounding box
        ax.set_xlim(6.5, 18.6)
        ax.set_ylim(36.5, 47.2)
        ax.set_aspect("equal")
        ax.set_title("Distribuzione Geografica Ricavi — Italia",
                     color=C["text"], pad=12)
        ax.set_xlabel("Longitudine", color=C["dim"])
        ax.set_ylabel("Latitudine", color=C["dim"])
        ax.grid(True, color=C["chart_grid"], linewidth=0.3, alpha=0.5)

        if not matched:
            ax.text(
                12.5, 42,
                "Nessuna città riconosciuta.\nVerifica i nomi (es. Roma, Milano, Napoli).",
                color=C["dim"], fontsize=9, ha="center", va="center",
                bbox=dict(boxstyle="round,pad=0.4", facecolor=C["card"], alpha=0.7),
            )
            self._canvas_widget.draw()
            return

        max_rev = max(v[2] for v in matched.values())
        min_rev = min(v[2] for v in matched.values())
        rev_range = max_rev - min_rev if max_rev != min_rev else 1

        for i, (city, (lat, lon, rev)) in enumerate(
            sorted(matched.items(), key=lambda x: x[1][2], reverse=True)
        ):
            size = 80 + 600 * (rev - min_rev) / rev_range
            color = CHART_COLORS[i % len(CHART_COLORS)]
            ax.scatter(lon, lat, s=size, c=color, alpha=0.75, zorder=3,
                       edgecolors="white", linewidths=0.4)
            ax.annotate(
                f"{city}\n€{rev:,.0f}",
                (lon, lat),
                textcoords="offset points", xytext=(6, 4),
                fontsize=6.5, color=C["text"],
                bbox=dict(boxstyle="round,pad=0.2", facecolor=C["card"], alpha=0.6),
            )

        if unmatched:
            ax.set_xlabel(
                f"Longitudine  |  Città non riconosciute ({len(unmatched)}): "
                + ", ".join(unmatched[:6]) + ("…" if len(unmatched) > 6 else ""),
                color=C["dim"], fontsize=7,
            )

        fig.tight_layout()
        self._canvas_widget.draw()


# ═══════════════════════════════════════════════════════════════════════════════
# Download Progress Dialog
# ═══════════════════════════════════════════════════════════════════════════════

class DownloadDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Download Aggiornamento")
        self.configure(bg=C["bg"])
        self.resizable(False, False)
        self.grab_set()
        self.geometry("420x160")

        self._center(parent)

        tk.Label(self, text="Download aggiornamento in corso…",
                 bg=C["bg"], fg=C["text"], font=("Segoe UI", 10)).pack(pady=(20, 8))

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Accent.Horizontal.TProgressbar",
                        troughcolor=C["card"],
                        background=C["accent"],
                        bordercolor=C["border"])
        self._pb = ttk.Progressbar(self, style="Accent.Horizontal.TProgressbar",
                                    length=380, mode="determinate")
        self._pb.pack(pady=4)
        self._lbl = tk.Label(self, text="0 / ? MB", bg=C["bg"], fg=C["dim"],
                              font=("Segoe UI", 9))
        self._lbl.pack()

        self._retry_btn: Optional[tk.Button] = None

    def _center(self, parent):
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_rootx(), parent.winfo_rooty()
        self.geometry(f"+{px + (pw - 420) // 2}+{py + (ph - 160) // 2}")

    def update_progress(self, downloaded: int, total: int):
        def _do():
            mb_down = downloaded / 1_048_576
            mb_tot = total / 1_048_576 if total else 0
            self._lbl.config(text=f"{mb_down:.1f} / {mb_tot:.1f} MB")
            if total:
                self._pb["value"] = downloaded / total * 100
            else:
                self._pb["mode"] = "indeterminate"
                self._pb.start()
        self.after(0, _do)

    def show_error(self, msg: str, on_retry):
        def _do():
            self._lbl.config(text=f"Errore: {msg}", fg=C["error"])
            if self._retry_btn:
                self._retry_btn.destroy()
            self._retry_btn = tk.Button(
                self, text="Riprova", bg=C["accent"], fg=C["text"],
                relief="flat", padx=12, pady=4,
                command=on_retry,
            )
            self._retry_btn.pack(pady=8)
        self.after(0, _do)

    def show_complete(self):
        def _do():
            self._pb["value"] = 100
            self._lbl.config(
                text="Aggiornamento scaricato. Riavvia l'app per applicarlo.",
                fg=C["success"],
            )
        self.after(0, _do)


# ═══════════════════════════════════════════════════════════════════════════════
# Main Application
# ═══════════════════════════════════════════════════════════════════════════════

class SalesAnalyzerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Sales Analyzer v{APP_VERSION}")
        self.configure(bg=C["bg"])
        self.minsize(1000, 640)
        self.geometry("1280x780")

        try:
            self.iconbitmap(str(BASE_DIR / "icon.ico"))
        except Exception:
            pass

        self._config = ConfigManager()
        self._mapper = ColumnMapper()

        self._df_raw: Optional[pd.DataFrame] = None
        self._df_filtered: Optional[pd.DataFrame] = None
        self._mapping: Optional[Dict[str, Optional[str]]] = None
        self._current_view: str = "trend"
        self._update_banner_dismissed = False
        self._update_banner_widget: Optional[UpdateBanner] = None

        self._style_ttk()
        self._build_ui()
        self._check_updates_async()

    # ── TTK styling ────────────────────────────────────────────────────────────
    def _style_ttk(self):
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("TCombobox",
                        fieldbackground=C["card"],
                        background=C["card"],
                        foreground=C["text"],
                        bordercolor=C["border"],
                        arrowcolor=C["text"])
        style.map("TCombobox", fieldbackground=[("readonly", C["card"])])
        style.configure("TSeparator", background=C["border"])

    # ── UI layout ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        # Outer container (banner + rest)
        self._outer = tk.Frame(self, bg=C["bg"])
        self._outer.pack(fill="both", expand=True)

        # Toolbar
        self._build_toolbar()

        # Separator
        ttk.Separator(self._outer, orient="horizontal").pack(fill="x")

        # Body: sidebar + main area
        body = tk.Frame(self._outer, bg=C["bg"])
        body.pack(fill="both", expand=True)

        self._build_sidebar(body)

        sep = tk.Frame(body, bg=C["border"], width=1)
        sep.pack(side="left", fill="y")

        self._main_area = tk.Frame(body, bg=C["bg"])
        self._main_area.pack(side="left", fill="both", expand=True)

        self._build_views()
        self._switch_view("trend")

    def _build_toolbar(self):
        tb = tk.Frame(self._outer, bg=C["sidebar"], pady=6)
        tb.pack(fill="x")

        # Load button
        self._load_btn = tk.Button(
            tb, text="📂 Carica Excel",
            bg=C["accent"], fg=C["text"],
            relief="flat", padx=14, pady=6,
            font=("Segoe UI", 10, "bold"),
            cursor="hand2",
            activebackground=C["accent2"], activeforeground=C["text"],
            command=self._load_excel,
        )
        self._load_btn.pack(side="left", padx=(12, 6))

        sep1 = tk.Frame(tb, bg=C["border"], width=1)
        sep1.pack(side="left", fill="y", padx=8, pady=4)

        # Date range
        tk.Label(tb, text="Da:", bg=C["sidebar"], fg=C["dim"],
                 font=("Segoe UI", 9)).pack(side="left")
        self._date_from = self._make_date_widget(tb)
        self._date_from.pack(side="left", padx=(2, 8))

        tk.Label(tb, text="A:", bg=C["sidebar"], fg=C["dim"],
                 font=("Segoe UI", 9)).pack(side="left")
        self._date_to = self._make_date_widget(tb)
        self._date_to.pack(side="left", padx=2)

        sep2 = tk.Frame(tb, bg=C["border"], width=1)
        sep2.pack(side="left", fill="y", padx=8, pady=4)

        # Channel filter
        self._channel_filter = MultiSelectDropdown(tb, label="Canale", bg=C["sidebar"])
        self._channel_filter.pack(side="left", padx=4)

        sep3 = tk.Frame(tb, bg=C["border"], width=1)
        sep3.pack(side="left", fill="y", padx=8, pady=4)

        # Apply filters button
        self._apply_btn = tk.Button(
            tb, text="🔍 Applica Filtri",
            bg=C["card"], fg=C["text"],
            relief="flat", padx=12, pady=6,
            font=("Segoe UI", 9),
            cursor="hand2",
            activebackground=C["hover"], activeforeground=C["text"],
            command=self._apply_filters,
        )
        self._apply_btn.pack(side="left", padx=4)

        # Updates button on right
        self._update_btn = tk.Button(
            tb, text="🔄 Aggiornamenti",
            bg=C["card"], fg=C["dim"],
            relief="flat", padx=12, pady=6,
            font=("Segoe UI", 9),
            cursor="hand2",
            activebackground=C["hover"], activeforeground=C["text"],
            command=self._check_updates_manual,
        )
        self._update_btn.pack(side="right", padx=12)

        self._filename_lbl = tk.Label(
            tb, text="Nessun file caricato",
            bg=C["sidebar"], fg=C["dim"],
            font=("Segoe UI", 9, "italic"),
        )
        self._filename_lbl.pack(side="right", padx=8)

    def _make_date_widget(self, parent) -> tk.Widget:
        if HAS_CALENDAR:
            from tkcalendar import DateEntry
            de = DateEntry(parent, width=10, date_pattern="dd/MM/yyyy",
                           background=C["card"], foreground=C["text"],
                           bordercolor=C["border"], headersbackground=C["card"],
                           headersforeground=C["text"], selectbackground=C["accent"],
                           normalbackground=C["bg"], normalforeground=C["text"],
                           weekendbackground=C["bg"], weekendforeground=C["dim"],
                           font=("Segoe UI", 9))
            return de
        else:
            e = tk.Entry(parent, width=11, bg=C["card"], fg=C["text"],
                         insertbackground=C["text"],
                         relief="flat", font=("Segoe UI", 9))
            e.insert(0, "GG/MM/AAAA")
            e.bind("<FocusIn>", lambda ev: (
                e.delete(0, "end") if e.get() in ("GG/MM/AAAA", "") else None
            ))
            return e

    def _get_date_value(self, widget) -> Optional[date]:
        if HAS_CALENDAR:
            try:
                return widget.get_date()
            except Exception:
                return None
        else:
            val = widget.get().strip()
            if not val or val == "GG/MM/AAAA":
                return None
            for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d"):
                try:
                    return datetime.strptime(val, fmt).date()
                except ValueError:
                    continue
            return None

    def _build_sidebar(self, parent):
        sb = tk.Frame(parent, bg=C["sidebar"], width=170)
        sb.pack(side="left", fill="y")
        sb.pack_propagate(False)

        tk.Label(sb, text="VISTE",
                 bg=C["sidebar"], fg=C["dim"],
                 font=("Segoe UI", 8, "bold")).pack(pady=(20, 4), padx=16, anchor="w")

        self._nav_btns: Dict[str, tk.Button] = {}
        views = [
            ("trend",       "📈  Trend"),
            ("top",         "🏆  Top Prodotti"),
            ("comparison",  "📊  Confronto Periodi"),
            ("geo",         "🗺  Mappa Geo"),
        ]
        for key, label in views:
            btn = tk.Button(
                sb, text=label,
                bg=C["sidebar"], fg=C["text"],
                relief="flat", anchor="w",
                padx=16, pady=10,
                font=("Segoe UI", 10),
                cursor="hand2",
                activebackground=C["hover"], activeforeground=C["text"],
                command=lambda k=key: self._switch_view(k),
            )
            btn.pack(fill="x")
            self._nav_btns[key] = btn

        # Footer
        tk.Label(sb, text=f"v{APP_VERSION}",
                 bg=C["sidebar"], fg=C["dim"],
                 font=("Segoe UI", 8)).pack(side="bottom", pady=8)

    def _build_views(self):
        self._views: Dict[str, BaseView] = {
            "trend":      TrendView(self._main_area, self),
            "top":        TopProductsView(self._main_area, self),
            "comparison": PeriodComparisonView(self._main_area, self),
            "geo":        GeoMapView(self._main_area, self),
        }
        for v in self._views.values():
            v.place(relx=0, rely=0, relwidth=1, relheight=1)

    def _switch_view(self, key: str):
        self._current_view = key
        for k, v in self._views.items():
            if k == key:
                v.lift()
            else:
                v.lower()

        for k, btn in self._nav_btns.items():
            btn.config(
                bg=C["accent"] if k == key else C["sidebar"],
                fg=C["text"],
            )

        if self._df_filtered is not None and self._mapping:
            self._views[key].refresh(self._df_filtered, self._mapping)

    # ── Excel loading ──────────────────────────────────────────────────────────
    def _load_excel(self):
        initial_dir = self._config.get("last_folder", str(Path.home()))
        path = filedialog.askopenfilename(
            title="Apri file Excel",
            initialdir=initial_dir,
            filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm"), ("All Files", "*.*")],
        )
        if not path:
            return

        self._config.set("last_folder", str(Path(path).parent))
        self._show_loading()

        def _load():
            try:
                df = pd.read_excel(path, engine="openpyxl")
                self.after(0, lambda: self._on_data_loaded(df, Path(path).name))
            except Exception as exc:
                log.error(f"Excel load error: {exc}")
                self.after(0, lambda: self._on_load_error(str(exc)))

        threading.Thread(target=_load, daemon=True).start()

    def _show_loading(self):
        self._loading_overlay = LoadingOverlay(self._main_area)
        self._loading_overlay.place(relx=0, rely=0, relwidth=1, relheight=1)
        self._loading_overlay.lift()

    def _hide_loading(self):
        try:
            self._loading_overlay.destroy()
        except Exception:
            pass

    def _on_data_loaded(self, df: pd.DataFrame, filename: str):
        self._hide_loading()

        # Parse dates in all object columns
        for col in df.columns:
            if df[col].dtype == object:
                try:
                    df[col] = pd.to_datetime(df[col], dayfirst=True, infer_datetime_format=True)
                except Exception:
                    pass

        cols = list(df.columns)
        mapping = self._mapper.auto_map(cols)

        if not self._mapper.is_complete(mapping):
            dlg = ColumnMappingDialog(self, cols, mapping)
            self.wait_window(dlg)
            if dlg.result is None:
                return
            mapping = dlg.result
        else:
            # Check ambiguous — show mapping UI if any required field matched multiple
            ambiguous = any(mapping.get(r) is None for r in ("quantity", "channel", "geo"))
            if not self._mapper.is_complete(mapping):
                dlg = ColumnMappingDialog(self, cols, mapping)
                self.wait_window(dlg)
                if dlg.result is None:
                    return
                mapping = dlg.result

        # Convert date column
        date_col = mapping["date"]
        try:
            df[date_col] = pd.to_datetime(df[date_col], dayfirst=True)
        except Exception as e:
            messagebox.showerror("Errore Colonna Data",
                                 f"Impossibile interpretare la colonna data:\n{e}")
            return

        # Convert revenue and quantity columns
        for role in ("revenue", "quantity"):
            col = mapping.get(role)
            if col:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        self._df_raw = df
        self._mapping = mapping
        self._df_filtered = df.copy()

        # Populate channel filter
        channel_col = mapping.get("channel")
        if channel_col and channel_col in df.columns:
            self._channel_filter.set_options(df[channel_col].dropna().unique().tolist())

        # Populate date range defaults
        date_min = df[date_col].min().date()
        date_max = df[date_col].max().date()
        self._set_date_widget(self._date_from, date_min)
        self._set_date_widget(self._date_to, date_max)

        self.title(f"Sales Analyzer v{APP_VERSION} — {filename}")
        self._filename_lbl.config(text=filename, fg=C["text"])

        self._switch_view(self._current_view)

    def _set_date_widget(self, widget, d: date):
        if HAS_CALENDAR:
            try:
                widget.set_date(d)
            except Exception:
                pass
        else:
            widget.delete(0, "end")
            widget.insert(0, d.strftime("%d/%m/%Y"))

    def _on_load_error(self, msg: str):
        self._hide_loading()
        messagebox.showerror(
            "Errore Caricamento",
            f"Impossibile aprire il file Excel.\n\n{msg}\n\n"
            "Assicurati che:\n"
            "  • Il file sia un .xlsx valido\n"
            "  • Non sia aperto in Excel\n"
            "  • Contenga almeno una colonna data, una prodotto e una ricavo",
        )

    # ── Filters ────────────────────────────────────────────────────────────────
    def _apply_filters(self):
        if self._df_raw is None:
            messagebox.showinfo("Filtri", "Carica prima un file Excel.")
            return

        df = self._df_raw.copy()
        date_col = self._mapping["date"]
        channel_col = self._mapping.get("channel")

        date_from = self._get_date_value(self._date_from)
        date_to = self._get_date_value(self._date_to)

        if date_from:
            df = df[df[date_col].dt.date >= date_from]
        if date_to:
            df = df[df[date_col].dt.date <= date_to]

        if channel_col and channel_col in df.columns:
            selected = self._channel_filter.get_selected()
            if selected:
                df = df[df[channel_col].isin(selected)]

        if df.empty:
            messagebox.showwarning("Filtri",
                                   "Nessun dato corrisponde ai filtri selezionati.")
            return

        self._df_filtered = df
        self._views[self._current_view].refresh(df, self._mapping)

    # ── Auto-updater ───────────────────────────────────────────────────────────
    def _check_updates_async(self):
        try:
            from updater import check_for_update
            check_for_update(
                on_update_available=self._on_update_available,
                on_error=lambda e: log.info(f"Update check silently failed: {e}"),
            )
        except Exception as e:
            log.info(f"Updater not available: {e}")

    def _check_updates_manual(self):
        self._update_btn.config(text="Controllo…", state="disabled")
        try:
            from updater import check_for_update
            def on_available(ver, url):
                self.after(0, lambda: self._on_update_available(ver, url))
                self.after(0, lambda: self._update_btn.config(
                    text="🔄 Aggiornamenti", state="normal"))

            def on_error(e):
                self.after(0, lambda: self._update_btn.config(
                    text="🔄 Aggiornamenti", state="normal"))
                self.after(0, lambda: messagebox.showinfo(
                    "Aggiornamenti", f"Nessun aggiornamento disponibile.\n({e})"))

            def on_up_to_date():
                self.after(0, lambda: self._update_btn.config(
                    text="🔄 Aggiornamenti", state="normal"))
                self.after(0, lambda: messagebox.showinfo(
                    "Aggiornamenti", f"Sei già alla versione più recente (v{APP_VERSION})."))

            import threading
            from updater import get_local_version, GITHUB_API
            import requests
            from packaging.version import Version

            def _check():
                try:
                    resp = requests.get(GITHUB_API, timeout=8)
                    resp.raise_for_status()
                    data = resp.json()
                    tag = data.get("tag_name", "").lstrip("v").strip()
                    if tag and Version(tag) > Version(get_local_version()):
                        assets = data.get("assets", [])
                        url = next((a["browser_download_url"] for a in assets
                                    if a.get("name", "").lower().endswith(".exe")), "")
                        on_available(tag, url)
                    else:
                        on_up_to_date()
                except Exception as exc:
                    on_error(str(exc))

            threading.Thread(target=_check, daemon=True).start()
        except Exception as e:
            self._update_btn.config(text="🔄 Aggiornamenti", state="normal")
            messagebox.showinfo("Aggiornamenti",
                                f"Verifica aggiornamenti non disponibile.\n({e})")

    def _on_update_available(self, version: str, download_url: str):
        if self._update_banner_dismissed or self._update_banner_widget:
            return

        def on_download():
            self._start_download(version, download_url)

        banner = UpdateBanner(self._outer, version=version, on_download=on_download)
        banner.pack(fill="x", before=self._outer.winfo_children()[0])
        self._update_banner_widget = banner

        original_destroy = banner.destroy
        def on_dismiss():
            self._update_banner_dismissed = True
            self._update_banner_widget = None
            original_destroy()
        banner.destroy = on_dismiss
        # Re-bind the ✕ button
        for child in banner.winfo_children():
            if isinstance(child, tk.Button) and child.cget("text") == "✕":
                child.config(command=on_dismiss)

    def _start_download(self, version: str, url: str):
        if not url:
            messagebox.showerror("Download",
                                 "URL di download non disponibile per questa release.")
            return

        dlg = DownloadDialog(self)

        def on_progress(downloaded, total):
            dlg.update_progress(downloaded, total)

        def on_complete(path):
            dlg.show_complete()
            log.info(f"Update downloaded to {path}")
            self.after(3000, lambda: _apply())

        def _apply():
            try:
                from updater import apply_update
                apply_update(path)
            except Exception as e:
                messagebox.showerror("Errore", f"Impossibile applicare l'aggiornamento:\n{e}")

        def on_error(msg):
            dlg.show_error(msg, on_retry=lambda: self._start_download(version, url))

        try:
            from updater import download_update
            download_update(url, on_progress, on_complete, on_error)
        except Exception as e:
            dlg.show_error(str(e), on_retry=lambda: self._start_download(version, url))


# ═══════════════════════════════════════════════════════════════════════════════
# Entry point
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    app = SalesAnalyzerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
