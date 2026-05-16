"""
Aplicación de Impresión de Etiquetas de Muebles
Navega por categorías desde Excel y genera etiquetas imprimibles
"""

import tkinter as tk
from tkinter import messagebox, filedialog
import pandas as pd
from datetime import date
import os
import re
import sys
import subprocess
import tempfile
import json
import unicodedata

try:
    from PIL import Image, ImageTk
    PIL_UI_AVAILABLE = True
except ImportError:
    Image = None
    ImageTk = None
    PIL_UI_AVAILABLE = False

# Impresión directa Windows (pywin32) gkhk
try:
    import win32print
    import win32api
    WIN32_AVAILABLE = True
except ImportError:
    win32print = None
    win32api = None
    WIN32_AVAILABLE = False

# Fichero para recordar la última impresora seleccionada
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PRINTER_CONFIG = os.path.join(BASE_DIR, ".last_printer.json")

def save_last_printer(name: str):
    try:
        with open(PRINTER_CONFIG, "w", encoding="utf-8") as f:
            json.dump({"printer": name}, f)
    except Exception:
        pass

def load_last_printer() -> str:
    try:
        with open(PRINTER_CONFIG, "r", encoding="utf-8") as f:
            return json.load(f).get("printer", "")
    except Exception:
        return ""

# ─────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────
EXCEL_FILE = r"M:\01. BBDD\01. ETIQUETADORA\productos.xlsx"
ASSETS_DIR = os.path.join(BASE_DIR, "assets")

# Colores y tipografías
BG_DARK       = "#1C1C1C"  # Very dark gray
BG_CARD       = "#2E2E2E"  # Dark gray
BG_HOVER      = "#404040"  # Medium dark gray
ACCENT        = "#606060"  # Medium gray
ACCENT2       = "#808080"  # Light medium gray
TEXT_WHITE    = "#FFFFFF"  # White
TEXT_GRAY     = "#B0B0B0"  # Light gray
BTN_BACK      = "#505050"  # Medium gray

FONT_TITLE    = ("Georgia", 22, "bold")
FONT_SUBTITLE = ("Georgia", 13)
FONT_BTN      = ("Segoe UI", 13, "bold")
FONT_BTN_SM   = ("Segoe UI", 10)
FONT_CRUMB    = ("Segoe UI", 9)
FONT_COUNTER  = ("Georgia", 28, "bold")
SCROLLBAR_WIDTH = 28
# Tamaño de las miniaturas (mayor = ocupa más espacio dentro del botón)
LEVEL_IMAGE_SIZE_BIG = (140, 140)
LEVEL_IMAGE_SIZE_SMALL = (105, 105)

# Emojis por categoría (nivel 1)
CATEGORY_ICONS = {
   
}

def get_icon(name: str) -> str:
    key = name.lower().strip()
    for k, v in CATEGORY_ICONS.items():
        if k in key:
            return v
    return ""


# ─────────────────────────────────────────────
# CARGA DE DATOS DESDE EXCEL
# ─────────────────────────────────────────────
def normalize_asset_key(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", str(value or ""))
    ascii_text = normalized.encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]+", "_", ascii_text.lower()).strip("_")


def get_level_asset_dir(level_num: int) -> str:
    return os.path.join(ASSETS_DIR, f"nivel_{level_num}")


def ensure_level_asset_dirs(level_count: int):
    os.makedirs(ASSETS_DIR, exist_ok=True)
    for level_num in range(1, level_count + 1):
        os.makedirs(get_level_asset_dir(level_num), exist_ok=True)


def load_excel(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, dtype=str)
    df.fillna("", inplace=True)
    # Normalizar: columnas nivel_1, nivel_2, ... nivel_N
    return df


def get_level_columns(df: pd.DataFrame) -> list:
    """Devuelve columnas nivel_1, nivel_2, ... en orden"""
    cols = [c for c in df.columns if c.lower().startswith("nivel_")]
    cols.sort(key=lambda x: int(x.split("_")[1]))
    return cols


def get_options_at_level(df: pd.DataFrame, level_cols: list,
                         selections: list) -> list:
    """Filtra df según selecciones previas y devuelve opciones únicas del siguiente nivel"""
    filtered = df.copy()
    for i, sel in enumerate(selections):
        if i < len(level_cols):
            filtered = filtered[filtered[level_cols[i]].str.strip().str.lower()
                                == sel.strip().lower()]
    next_level = len(selections)
    if next_level >= len(level_cols):
        return []
    options = filtered[level_cols[next_level]].str.strip()
    options = options[options != ""].unique().tolist()
    return sorted(options)


def get_product_row(df: pd.DataFrame, level_cols: list,
                    selections: list) -> pd.Series | None:
    filtered = df.copy()
    for i, sel in enumerate(selections):
        if i < len(level_cols):
            filtered = filtered[filtered[level_cols[i]].str.strip().str.lower()
                                == sel.strip().lower()]
    if len(filtered) == 1:
        return filtered.iloc[0]
    elif len(filtered) > 1:
        return filtered.iloc[0]
    return None


# ─────────────────────────────────────────────
# WIDGETS REUTILIZABLES
# ─────────────────────────────────────────────
class HoverButton(tk.Frame):
    def __init__(self, parent, text, icon="", image=None, command=None,
                 big=False, color=BG_CARD, **kwargs):
        super().__init__(parent, bg=BG_DARK, **kwargs)
        self.command = command
        self.color = color
        self.hover_color = BG_HOVER
        self.image_ref = image

        # Reducimos el padding para que, al aumentar la miniatura, el botón no crezca demasiado
        pad = 18 if big else 12
        fsize = 15 if big else 13

        self.inner = tk.Frame(self, bg=color, cursor="hand2",
                              padx=pad, pady=pad)
        self.inner.pack(fill="both", expand=True)

        has_visual = image is not None or bool(icon)

        if image is not None:
            tk.Label(self.inner, image=image, bg=color).pack(pady=(0, 8))
        elif icon:
            tk.Label(self.inner, text=icon, bg=color,
                     font=("Segoe UI Emoji", 28 if big else 20),
                     fg=TEXT_WHITE).pack(pady=(0, 6))

        tk.Label(self.inner, text=text, bg=color,
                 font=(FONT_BTN[0], fsize, "bold"),
                 fg=TEXT_WHITE,
                 justify="center",
                 wraplength=160 if has_visual else 220).pack(
                     fill="x", expand=not has_visual
                 )

        for w in [self.inner] + self.inner.winfo_children():
            w.bind("<Enter>", self._on_enter)
            w.bind("<Leave>", self._on_leave)
            w.bind("<Button-1>", self._on_click)

    def _on_enter(self, _=None):
        self._set_color(self.hover_color)

    def _on_leave(self, _=None):
        self._set_color(self.color)

    def _set_color(self, c):
        self.inner.config(bg=c)
        for w in self.inner.winfo_children():
            w.config(bg=c)

    def _on_click(self, _=None):
        if self.command:
            self.command()


# ─────────────────────────────────────────────
# VENTANA PRINCIPAL
# ─────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Etiquetas de Mercancías - Cancio Fábrica de Muebles, S.A.")
        self.geometry("900x640")
        self.resizable(True, True)
        # Siempre iniciar maximizado en Windows (y en plataformas que lo soporten)
        try:
            self.state("zoomed")
        except Exception:
            pass
        self.configure(bg=BG_DARK)

        # Estado
        self.df = None
        self.level_cols = []
        self.selections = []          # selecciones acumuladas
        self.excel_path = EXCEL_FILE
        self.asset_index_cache = {}
        self.option_image_cache = {}

        # ── Barra de impresora (parte inferior, siempre visible) ──
        self.printer_bar = tk.Frame(self, bg=BG_CARD, pady=8)
        self.printer_bar.pack(side="bottom", fill="x")

        tk.Label(self.printer_bar, text="Impresora:",
                 bg=BG_CARD, fg=TEXT_GRAY,
                 font=FONT_BTN_SM).pack(side="left", padx=(14, 6))

        self.printer_var = tk.StringVar()
        self.printer_menu = None
        self._build_printer_menu()

        # Logo header
        self.logo_frame = tk.Frame(self, bg="#333333")
        self.logo_frame.pack(side="top", fill="x")
        self.company_label = tk.Label(self.logo_frame, text="CANCIO Fábrica de Muebles, S.A.", bg="#333333", fg="white", font=("Georgia", 20, "bold"))
        self.time_label = tk.Label(self.logo_frame, text="", bg="#333333", fg="white", font=("Georgia", 16))
        self.company_label.pack(side="left", padx=20)
        self.time_label.pack(side="right", padx=20)
        self._update_time()

        self.container = tk.Frame(self, bg=BG_DARK)
        self.container.pack(fill="both", expand=False)

        if os.path.exists(self.excel_path):
            self._load_data(self.excel_path)
        else:
            self._show_load_screen()


    def _update_time(self):
        from datetime import datetime
        now = datetime.now()
        self.time_label.config(text=now.strftime("%d/%m/%Y %H:%M:%S"))
        self.after(1000, self._update_time)

    def _direct_print(self):
        code = self.ref_entry.get().strip()
        if not code:
            messagebox.showinfo("Info", "Introduce un código referencia")
            return
        if self.df is None:
            messagebox.showerror("Error", "Datos no cargados")
            return
        filtered = self.df[self.df['referencia'].str.strip().str.lower() == code.lower()]
        if filtered.empty:
            messagebox.showerror("Error", f"Código '{code}' no encontrado")
            return
        row = filtered.iloc[0]
        # Build selections
        self.selections = []
        for col in self.level_cols:
            val = row[col].strip()
            if val:
                self.selections.append(val)
        # Now show print screen
        self._show_print_screen()
        self.ref_entry.delete(0, tk.END)


    def _reset_asset_cache(self):
        self.asset_index_cache = {}
        self.option_image_cache = {}

    def _get_level_asset_index(self, level_num: int) -> dict:
        cached = self.asset_index_cache.get(level_num)
        if cached is not None:
            return cached

        level_dir = get_level_asset_dir(level_num)
        index = {}
        if os.path.isdir(level_dir):
            for root, dirnames, files in os.walk(level_dir):
                dirnames.sort()
                for filename in sorted(files):
                    if not filename.lower().endswith(".png"):
                        continue
                    rel_path = os.path.relpath(os.path.join(root, filename), level_dir)
                    rel_parts = os.path.splitext(rel_path)[0].replace("\\", "/").split("/")
                    key = tuple(
                        part for part in (normalize_asset_key(item) for item in rel_parts)
                        if part
                    )
                    if key and key not in index:
                        index[key] = os.path.join(root, filename)

        self.asset_index_cache[level_num] = index
        return index

    def _get_option_image_path(self, level_num: int, option: str) -> str | None:
        option_key = normalize_asset_key(option)
        if not option_key:
            return None

        index = self._get_level_asset_index(level_num)
        context_keys = [normalize_asset_key(sel) for sel in self.selections]
        context_keys = [key for key in context_keys if key]

        candidates = []
        if context_keys:
            candidates.append(tuple(context_keys + [option_key]))
        candidates.append((option_key,))

        for candidate in candidates:
            image_path = index.get(candidate)
            if image_path:
                return image_path
        return None

    def _get_option_image(self, level_num: int, option: str, big: bool):
        if not PIL_UI_AVAILABLE:
            return None

        image_path = self._get_option_image_path(level_num, option)
        if not image_path:
            return None

        size = LEVEL_IMAGE_SIZE_BIG if big else LEVEL_IMAGE_SIZE_SMALL
        cache_key = (image_path, size)
        cached = self.option_image_cache.get(cache_key)
        if cached is not None:
            return cached

        try:
            with Image.open(image_path) as source_image:
                prepared = source_image.convert("RGBA")
            prepared.thumbnail(size, Image.LANCZOS)

            canvas = Image.new("RGBA", size, (0, 0, 0, 0))
            pos_x = (size[0] - prepared.width) // 2
            pos_y = (size[1] - prepared.height) // 2
            canvas.paste(prepared, (pos_x, pos_y), prepared)

            photo = ImageTk.PhotoImage(canvas)
            self.option_image_cache[cache_key] = photo
            return photo
        except Exception:
            return None

    # ── Menú de impresoras ──────────────────────
    def _build_printer_menu(self):
        if WIN32_AVAILABLE:
            try:
                printers = [p[2] for p in win32print.EnumPrinters(
                    win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
            except Exception:
                printers = []
        else:
            printers = []

        if not printers:
            tk.Label(self.printer_bar,
                     text="(instala pywin32 para imprimir directamente)",
                     bg=BG_CARD, fg=ACCENT, font=FONT_BTN_SM).pack(side="left")
            return

        # Seleccionar la última usada, o la predeterminada del sistema
        last = load_last_printer()
        default_sys = ""
        if WIN32_AVAILABLE:
            try:
                default_sys = win32print.GetDefaultPrinter()
            except Exception:
                pass

        if last in printers:
            self.printer_var.set(last)
        elif default_sys in printers:
            self.printer_var.set(default_sys)
        else:
            self.printer_var.set(printers[0])

        # Guardar al cambiar
        def on_change(*_):
            save_last_printer(self.printer_var.get())
        self.printer_var.trace_add("write", on_change)

        # OptionMenu (botón) acepta highlightthickness; su menú interno NO
        btn_kw  = dict(bg=BG_HOVER, fg=TEXT_WHITE,
                       activebackground=ACCENT, activeforeground=TEXT_WHITE,
                       relief="flat", font=FONT_BTN_SM, bd=0,
                       highlightthickness=0)
        menu_kw = dict(bg=BG_HOVER, fg=TEXT_WHITE,
                       activebackground=ACCENT, activeforeground=TEXT_WHITE,
                       font=FONT_BTN_SM, bd=0)
        menu = tk.OptionMenu(self.printer_bar, self.printer_var, *printers)
        menu.config(**btn_kw, padx=10, pady=4, cursor="hand2")
        menu["menu"].config(**menu_kw)
        menu.pack(side="left", padx=(0, 10))

        # Botón refrescar lista
        tk.Button(self.printer_bar, text="↺",
                  bg=BG_CARD, fg=TEXT_GRAY,
                  font=("Segoe UI", 11), relief="flat",
                  cursor="hand2", padx=4,
                  command=self._refresh_printers).pack(side="left")

    def _refresh_printers(self):
        for w in self.printer_bar.winfo_children():
            w.destroy()
        tk.Label(self.printer_bar, text="Impresora:",
                 bg=BG_CARD, fg=TEXT_GRAY,
                 font=FONT_BTN_SM).pack(side="left", padx=(14, 6))
        self._build_printer_menu()

    # ── Carga de datos ──────────────────────────
    def _load_data(self, path):
        try:
            self.df = load_excel(path)
            self.level_cols = get_level_columns(self.df)
            if not self.level_cols:
                messagebox.showerror("Error",
                    "El Excel no tiene columnas 'nivel_1', 'nivel_2', ...")
                self._show_load_screen()
                return
            self.excel_path = path
            self.selections = []
            ensure_level_asset_dirs(len(self.level_cols))
            self._reset_asset_cache()
            self._show_selection_screen()
        except Exception as e:
            messagebox.showerror("Error al leer Excel", str(e))
            self._show_load_screen()

    def _clear(self):
        for w in self.container.winfo_children():
            w.destroy()

    # ── Pantalla: cargar Excel ───────────────────
    def _show_load_screen(self):
        self._clear()
        f = tk.Frame(self.container, bg=BG_DARK)
        f.place(relx=0.5, rely=0.5, anchor="center")

        tk.Label(f, text="🗂️", bg=BG_DARK,
                 font=("Segoe UI Emoji", 48)).pack(pady=(0, 10))
        tk.Label(f, text="No se encontró productos.xlsx",
                 bg=BG_DARK, fg=TEXT_WHITE,
                 font=FONT_TITLE).pack()
        tk.Label(f, text="Carga tu archivo Excel de productos",
                 bg=BG_DARK, fg=TEXT_GRAY,
                 font=FONT_SUBTITLE).pack(pady=(4, 20))

        btn = tk.Button(f, text="📂  Seleccionar Excel",
                        bg=ACCENT, fg=TEXT_WHITE,
                        font=FONT_BTN, relief="flat",
                        padx=20, pady=12, cursor="hand2",
                        command=self._browse_excel)
        btn.pack()

    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="Selecciona el Excel de productos",
            filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self._load_data(path)

    # ── Pantalla: selección de opciones ─────────
    def _show_selection_screen(self):
        self._clear()
        options = get_options_at_level(self.df, self.level_cols, self.selections)

        # Si no hay más opciones → pantalla de impresión
        if not options:
            self._show_print_screen()
            return

        level_num = len(self.selections) + 1
        level_name = self.level_cols[len(self.selections)].replace("_", " ").title()

        # ── Header ──────────────────────────────
        header = tk.Frame(self.container, bg=BG_CARD, pady=16)
        header.pack(fill="x")

        crumb_parts = self.selections
        crumb_text = " › ".join(crumb_parts)

        # Título del nivel
        tk.Label(header,
                 text=crumb_text,
                 bg=BG_CARD, fg=TEXT_WHITE,
                 font=FONT_TITLE).pack(pady=(8, 0))
        tk.Label(header,
                 text=f"{len(options)} Opcion{'es' if len(options) != 1 else ''} Disponible{'s' if len(options) != 1 else ''}",
                 bg=BG_CARD, fg=TEXT_GRAY,
                 font=FONT_SUBTITLE).pack()

        # Buscador de referencia
        search_frame = tk.Frame(header, bg=BG_CARD)
        search_frame.pack(fill="x", pady=(10, 0))
        self.ref_label = tk.Label(search_frame, text="Código referencia:", bg=BG_CARD, fg=TEXT_GRAY, font=("Georgia", 20, "bold"))
        self.ref_entry = tk.Entry(search_frame, width=20, font=("Georgia", 20))
        self.ref_button = tk.Button(search_frame, text="Buscar", bg=ACCENT, fg=TEXT_WHITE, font=("Georgia", 20, "bold"), command=self._direct_print)
        self.ref_label.pack(side="left", padx=(0,5))
        self.ref_entry.pack(side="left", padx=(0,5))
        self.ref_button.pack(side="left")

        # ── Botón volver ────────────────────────
        if self.selections:
            back_bar = tk.Frame(self.container, bg=BG_DARK, pady=6)
            back_bar.pack(fill="x", padx=20)
            tk.Button(back_bar, text="← Volver",
                      bg=BTN_BACK, fg=TEXT_GRAY,
                      font=FONT_BTN_SM, relief="flat",
                      padx=10, pady=6, cursor="hand2",
                      command=self._go_back).pack(side="left")

        # Botón recargar Excel
        reload_bar = tk.Frame(self.container, bg=BG_DARK)
        reload_bar.pack(fill="x", padx=20)
        tk.Button(reload_bar, text="🔄 Recargar Excel",
                  bg=BG_DARK, fg=TEXT_GRAY,
                  font=FONT_BTN_SM, relief="flat",
                  padx=6, pady=4, cursor="hand2",
                  command=lambda: self._load_data(self.excel_path)).pack(side="right")

        # ── Grid de opciones ────────────────────
        canvas = tk.Canvas(self.container, bg=BG_DARK,
                           highlightthickness=0)
        scrollbar = tk.Scrollbar(self.container,
                                 orient="vertical",
                                 width=SCROLLBAR_WIDTH,
                                 command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(fill="both", expand=True, padx=20, pady=10)

        grid_frame = tk.Frame(canvas, bg=BG_DARK)
        canvas_window = canvas.create_window((0, 0),
                                             window=grid_frame,
                                             anchor="nw")

        def _on_frame_configure(_):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)

        grid_frame.bind("<Configure>", _on_frame_configure)
        canvas.bind("<Configure>", _on_canvas_configure)
        canvas.bind("<MouseWheel>",
                    lambda e: canvas.yview_scroll(-1*(e.delta//120), "units"))

        COLS = 4
        for idx, opt in enumerate(options):
            row, col = divmod(idx, COLS)
            image = self._get_option_image(level_num, opt, big=(level_num == 1))
            icon = get_icon(opt) if level_num == 1 else ""
            btn = HoverButton(grid_frame,
                              text=opt,
                              icon="" if image is not None else icon,
                              image=image,
                              big=(level_num == 1),
                              command=lambda o=opt: self._select(o))
            btn.grid(row=row, column=col,
                     padx=8, pady=8, sticky="nsew")

        for c in range(COLS):
            grid_frame.columnconfigure(c, weight=1)

    def _select(self, option: str):
        self.selections.append(option)
        self._show_selection_screen()

    def _go_back(self):
        if self.selections:
            self.selections.pop()
        self._show_selection_screen()

    # ── Pantalla: imprimir etiqueta ──────────────
    def _show_print_screen(self):
        self._clear()

        row = get_product_row(self.df, self.level_cols, self.selections)

        # ── Header ──────────────────────────────
        header = tk.Frame(self.container, bg=BG_CARD, pady=16)
        header.pack(fill="x")

        crumb_frame = tk.Frame(header, bg=BG_CARD)
        crumb_frame.pack()
        crumb_parts = self.selections
        for i, part in enumerate(crumb_parts):
            color = ACCENT if i == len(crumb_parts) - 1 else TEXT_GRAY
            tk.Label(crumb_frame, text=part,
                     bg=BG_CARD, fg=color,
                     font=FONT_CRUMB).pack(side="left")
            if i < len(crumb_parts) - 1:
                tk.Label(crumb_frame, text="  ›  ",
                         bg=BG_CARD, fg=TEXT_GRAY,
                         font=FONT_CRUMB).pack(side="left")

        tk.Label(header, text="🏷️  Imprimir Etiqueta",
                 bg=BG_CARD, fg=TEXT_WHITE,
                 font=FONT_TITLE).pack(pady=(8, 0))

        # ── Info del producto ────────────────────
        info_frame = tk.Frame(self.container, bg=BG_CARD,
                              padx=30, pady=20)
        info_frame.pack(fill="x", padx=20, pady=(10, 0))

        if row is not None:
            # Columnas que NO son niveles ni vacías
            data_cols = [c for c in self.df.columns
                         if not c.lower().startswith("nivel_")]

            for col in data_cols:
                val = str(row.get(col, "")).strip()
                if val:
                    r = tk.Frame(info_frame, bg=BG_CARD)
                    r.pack(fill="x", pady=2)
                    tk.Label(r, text=col.replace("_", " ").title() + ":",
                             bg=BG_CARD, fg=TEXT_GRAY,
                             font=("Segoe UI", 10),
                             width=18, anchor="e").pack(side="left")
                    tk.Label(r, text=val,
                             bg=BG_CARD, fg=TEXT_WHITE,
                             font=("Segoe UI", 11, "bold"),
                             anchor="w").pack(side="left", padx=(8, 0))
        else:
            tk.Label(info_frame,
                     text="⚠️  Producto no encontrado en el Excel",
                     bg=BG_CARD, fg=ACCENT,
                     font=FONT_SUBTITLE).pack()

        # ── Contador de etiquetas ────────────────
        counter_frame = tk.Frame(self.container, bg=BG_DARK, pady=20)
        counter_frame.pack()

        tk.Label(counter_frame, text="Cantidad de etiquetas",
                 bg=BG_DARK, fg=TEXT_GRAY,
                 font=FONT_SUBTITLE).pack()

        ctrl = tk.Frame(counter_frame, bg=BG_DARK)
        ctrl.pack(pady=8)

        self.qty_var = tk.IntVar(value=1)

        def change(delta):
            new_val = self.qty_var.get() + delta
            if new_val >= 1:
                self.qty_var.set(new_val)

        minus_btn = tk.Button(ctrl, text="−",
                              bg=ACCENT, fg=TEXT_WHITE,
                              font=("Georgia", 20, "bold"),
                              relief="flat", width=3, cursor="hand2",
                              command=lambda: change(-1))
        minus_btn.pack(side="left", padx=4)

        qty_label = tk.Label(ctrl, textvariable=self.qty_var,
                             bg=BG_DARK, fg=TEXT_WHITE,
                             font=FONT_COUNTER, width=4)
        qty_label.pack(side="left", padx=8)

        plus_btn = tk.Button(ctrl, text="+",
                             bg=ACCENT2, fg=TEXT_WHITE,
                             font=("Georgia", 20, "bold"),
                             relief="flat", width=3, cursor="hand2",
                             command=lambda: change(1))
        plus_btn.pack(side="left", padx=4)

        # ── Botones acción ───────────────────────
        actions = tk.Frame(self.container, bg=BG_DARK, pady=10)
        actions.pack()

        print_btn = tk.Button(actions,
                              text="🖨️   IMPRIMIR ETIQUETAS",
                              bg=ACCENT, fg=TEXT_WHITE,
                              font=("Segoe UI", 14, "bold"),
                              relief="flat", padx=30, pady=14,
                              cursor="hand2",
                              command=lambda: self._print_labels(row))
        print_btn.pack(pady=(0, 8))

        preview_btn = tk.Button(actions,
                                text="👁️  Vista previa (PDF)",
                                bg=BTN_BACK, fg=TEXT_WHITE,
                                font=FONT_BTN_SM,
                                relief="flat", padx=20, pady=8,
                                cursor="hand2",
                                command=lambda: self._preview_label(row))
        preview_btn.pack()

        # ── Volver ───────────────────────────────
        nav = tk.Frame(self.container, bg=BG_DARK, pady=4)
        nav.pack()
        tk.Button(nav, text="← Volver",
                  bg=BTN_BACK, fg=TEXT_GRAY,
                  font=FONT_BTN_SM, relief="flat",
                  padx=10, pady=6, cursor="hand2",
                  command=self._go_back).pack(side="left", padx=6)
        tk.Button(nav, text="🏠 INICIO",
                  bg=BTN_BACK, fg=TEXT_GRAY,
                  font=FONT_BTN_SM, relief="flat",
                  padx=10, pady=6, cursor="hand2",
                  command=self._restart).pack(side="left", padx=6)

    def _restart(self):
        self.selections = []
        self._show_selection_screen()

    # ── Generación de PDF de etiqueta ───────────
    def _build_pdf(self, row, qty: int, path: str):
        from reportlab.lib.units import cm
        from reportlab.pdfgen import canvas as rl_canvas
        from reportlab.lib import colors
        import io

        # Tamaño exacto según driver ZDesigner: 31cm × 7.62cm
        # El driver lo define como portrait (alto=310, ancho=76.2) pero el rollo
        # sale en landscape → PDF en landscape: W=310mm, H=76.2mm
        from reportlab.lib.units import mm
        W, H = 310 * mm, 76.2 * mm   # ancho × alto en landscape

        c = rl_canvas.Canvas(path, pagesize=(W, H))
        today = date.today().strftime("%d/%m/%Y")

        nombre     = str(row.get("nombre_producto", "")).strip() if row is not None else " ".join(self.selections)
        referencia = str(row.get("referencia", "")).strip()      if row is not None else ""
        codigo     = str(row.get("referencia", "")).strip()      if row is not None else ""

        # Intentar importar barcode (opcional)
        try:
            from barcode import Code128
            from barcode.writer import ImageWriter
            barcode_ok = True
        except ImportError:
            barcode_ok = False

        # Todas las cotas en puntos ReportLab (1mm = 2.8346 pt)
        # H = 76.2mm = 216 pt  → márgenes y posiciones ajustados a esa altura
        PAD_X  = 12          # margen lateral izquierdo/derecho
        PAD_Y  = 8           # margen vertical inferior
        CX     = W / 2       # centro horizontal

        for i in range(qty):
            if i > 0:
                c.showPage()

            # ── Fondo blanco ──────────────────────────
            c.setFillColor(colors.white)
            c.rect(0, 0, W, H, fill=1, stroke=0)

            # ── Borde rojo ────────────────────────────
            c.setStrokeColor(colors.black)
            c.setLineWidth(2)
            c.rect(2, 2, W - 4, H - 4, fill=0, stroke=1)

            # ── Breadcrumb (arriba centrado) ──────────
            crumb = " › ".join(self.selections)
            c.setFillColor(colors.black)
            c.setFont("Helvetica", 7)
            c.drawCentredString(CX, H - 14, crumb[:90])

            # ── Fecha (superior derecha) ───
            c.setFillColor(colors.black)
            c.setFont("Helvetica", 12)
            c.drawRightString(W - PAD_X, H - 30, today)

            # ── Nombre del producto (grande, centrado) ─
            # Fuente 32pt cabe bien en 76mm de alto
            FNAME, FSIZE = "Helvetica-Bold", 32
            c.setFillColor(colors.black)
            c.setFont(FNAME, FSIZE)
            max_w = W - PAD_X * 2
            words = nombre.split()
            lines, line = [], ""
            for w in words:
                test = (line + " " + w).strip()
                if c.stringWidth(test, FNAME, FSIZE) < max_w:
                    line = test
                else:
                    lines.append(line)
                    line = w
            if line:
                lines.append(line)
            lines = lines[::-1]  # Reverse the order of lines
            line_h   = FSIZE * 1.25        # interlineado
            total_h  = len(lines[:2]) * line_h
            start_y  = H / 2 - total_h / 2 + PAD_Y
            for li, txt in enumerate(lines[:2]):
                c.drawCentredString(CX, start_y + li * line_h, txt)

            # ── Referencia (abajo-izquierda, naranja) ─
            if referencia:
                c.setFillColor(colors.black)
                c.setFont("Helvetica-Bold", 12)
                c.drawString(PAD_X, PAD_Y + 8, f"Ref: {referencia}")

            # ── Código de barras CODE128 (abajo-derecha)
            if codigo and barcode_ok:
                try:
                    valor_bar = codigo.replace(" ", "") or "0"
                    barcode_obj = Code128(valor_bar, writer=ImageWriter())
                    buf = io.BytesIO()
                    barcode_obj.write(buf, options={
                        "write_text": False,
                        "module_width": 0.4,
                        "module_height": 10
                    })
                    buf.seek(0)
                    tmp_bc = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
                    tmp_bc.write(buf.getvalue())
                    tmp_bc.close()
                    bc_w, bc_h = 4.5 * cm, 1.4 * cm
                    bc_x = W - bc_w - PAD_X
                    bc_y = PAD_Y + 4
                    c.drawImage(tmp_bc.name, bc_x, bc_y, width=bc_w, height=bc_h)
                    c.setFillColor(colors.black)
                    c.setFont("Helvetica", 9)
                    c.drawCentredString(bc_x + bc_w / 2, bc_y - 9, valor_bar)
                except Exception:
                    c.setFillColor(colors.black)
                    c.setFont("Helvetica-Bold", 12)
                    c.drawRightString(W - PAD_X, PAD_Y + 8, codigo)
            elif codigo:
                c.setFillColor(colors.black)
                c.setFont("Helvetica-Bold", 12)
                c.drawRightString(W - PAD_X, PAD_Y + 8, codigo)

        c.save()

    def _preview_label(self, row):
        if row is None and not self.selections:
            messagebox.showwarning("Sin datos", "No hay producto seleccionado.")
            return
        try:
            tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
            tmp.close()
            self._build_pdf(row, self.qty_var.get(), tmp.name)
            # Abrir con visor PDF de Windows
            os.startfile(tmp.name)
        except Exception as e:
            messagebox.showerror("Error al generar PDF", str(e))

    def _send_to_printer(self, pdf_path: str, printer_name: str) -> bool:
        """
        Convierte el PDF a imagen PNG a la resolución del driver (203 dpi para Zebra)
        y lo envía directamente via win32print como bitmap GDI → escala 1:1 garantizada.
        Fallback: Sumatra con fit.
        """
        import shutil

        # ── 1. Impresión directa via GDI (win32print) ─────────────────────
        # Convierte cada página del PDF a imagen y la imprime pixel-perfect
        if WIN32_AVAILABLE:
            try:
                import fitz  # PyMuPDF
                import win32ui, win32con
                from PIL import Image

                # Abrir PDF y renderizar a 203 dpi (resolución Zebra ZDesigner)
                DPI = 203
                doc  = fitz.open(pdf_path)
                mat  = fitz.Matrix(DPI / 72, DPI / 72)

                hprinter = win32print.OpenPrinter(printer_name)
                try:
                    pinfo = win32print.GetPrinter(hprinter, 2)
                    hdc   = win32ui.CreateDC()
                    hdc.CreatePrinterDC(printer_name)
                    hdc.StartDoc(pdf_path)

                    for page in doc:
                        pix = page.get_pixmap(matrix=mat, alpha=False)
                        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)

                        # Ancho y alto imprimible en píxeles a la resolución del driver
                        pw = hdc.GetDeviceCaps(win32con.HORZRES)
                        ph = hdc.GetDeviceCaps(win32con.VERTRES)

                        # Escalar la imagen para que ocupe exactamente el área imprimible
                        img = img.resize((pw, ph), Image.LANCZOS)

                        # Convertir a BMP y enviar al DC
                        import io as _io
                        bmp_buf = _io.BytesIO()
                        img.save(bmp_buf, format="BMP")
                        bmp_buf.seek(0)

                        dib = win32ui.CreateBitmap()
                        dib.LoadBitmapData(bmp_buf.read(), pw, ph)

                        hdc.StartPage()
                        hdc.BitBlt((0, 0), (pw, ph), dib.GetHandle(), (0, 0), win32con.SRCCOPY)
                        hdc.EndPage()

                    hdc.EndDoc()
                    hdc.DeleteDC()
                finally:
                    win32print.ClosePrinter(hprinter)

                return True

            except Exception:
                pass  # si falla (fitz no instalado, etc.) caemos al método 2

        # ── 2. Sumatra PDF con -print-settings "fit" ──────────────────────
        # "fit" ajusta el PDF al papel configurado en el driver → mejor que noscale
        sumatra_paths = [
            r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
            r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "SumatraPDF", "SumatraPDF.exe"),
        ]
        for sp in sumatra_paths:
            if os.path.isfile(sp):
                subprocess.Popen([
                    sp,
                    "-print-to", printer_name,
                    "-print-settings", "fit",
                    "-silent",
                    "-exit-when-done",
                    pdf_path
                ], creationflags=0x08000000)
                return True

        return False

    def _print_labels(self, row):
        if row is None and not self.selections:
            messagebox.showwarning("Sin datos", "No hay producto seleccionado.")
            return
        try:
            tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
            tmp.close()
            self._build_pdf(row, self.qty_var.get(), tmp.name)

            printer_name = self.printer_var.get().strip() if WIN32_AVAILABLE else ""

            if printer_name:
                sent = self._send_to_printer(tmp.name, printer_name)
                if sent:
                    save_last_printer(printer_name)
                    messagebox.showinfo(
                        "Imprimiendo",
                        f"{self.qty_var.get()} etiqueta(s) enviadas a:\n{printer_name}"
                    )
                else:
                    # Ningún motor silencioso disponible → abrir diálogo normal
                    os.startfile(tmp.name)
                    messagebox.showinfo(
                        "Imprimiendo",
                        "Se abre el PDF para imprimir.\n\n"
                        "Para impresión directa sin diálogo instala SumatraPDF:\n"
                        "https://www.sumatrapdfreader.org"
                    )
            else:
                os.startfile(tmp.name)
        except Exception as e:
            messagebox.showerror("Error al imprimir", str(e))


# ─────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()
