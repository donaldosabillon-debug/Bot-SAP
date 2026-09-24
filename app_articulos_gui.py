"""
app_articulos_gui.py
Interfaz Gráfica (Tkinter) para el Bot de Actualización de Maestro de Artículos en SAP B1 HANA.
Incluye:
  - Ventana modal interactiva de Mapeo de Columnas Excel ➔ Campos SAP.
  - Calibración completa de 7 puntos con persistencia en JSON.
  - Botón de 'Probar con 1 artículo' para verificación antes de corridas masivas.
  - Ejecución continua segura con pausas y registro en vivo.
Lazarus & Lazarus.
"""

import ctypes
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2) # Per-monitor DPI aware
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

import os
import sys
import time
import threading
from typing import Optional
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import pyautogui

import window_manager
from bot_articulos_engine import BotArticulosEngine, SAPWindowNotFoundError


class ColumnMappingDialog(tk.Toplevel):
    """
    Ventana emergente interactiva para mapear las columnas del Excel cargado
    con los campos actualizables de SAP Business One.
    """
    def __init__(self, parent, df: pd.DataFrame, current_mapping: dict, current_flags: dict, on_apply_callback):
        super().__init__(parent)
        self.title("Mapeo de Columnas: Excel ➔ Campos de SAP B1")
        self.geometry("720x530")
        self.minsize(680, 480)
        self.transient(parent)
        self.grab_set()

        self.df = df
        self.on_apply = on_apply_callback
        self.columns_list = ["(Omitir / No actualizar)"] + list(df.columns)

        # Variables de selección de columnas
        self.var_item = tk.StringVar(value=current_mapping.get("item_col", ""))
        self.var_min = tk.StringVar(value=current_mapping.get("min_qty_col", "(Omitir / No actualizar)"))
        self.var_lead = tk.StringVar(value=current_mapping.get("lead_time_col", "(Omitir / No actualizar)"))
        self.var_tol = tk.StringVar(value=current_mapping.get("tolerance_col", "(Omitir / No actualizar)"))

        # Variables de activación de campos
        self.flag_min = tk.BooleanVar(value=current_flags.get("min_qty", True))
        self.flag_lead = tk.BooleanVar(value=current_flags.get("lead_time", True))
        self.flag_tol = tk.BooleanVar(value=current_flags.get("tolerance", True))

        # Variables para muestras en vivo
        self.sample_item = tk.StringVar()
        self.sample_min = tk.StringVar()
        self.sample_lead = tk.StringVar()
        self.sample_tol = tk.StringVar()

        self._build_ui()
        self._update_all_samples()

    def _build_ui(self):
        pad_frame = ttk.Frame(self, padding="15")
        pad_frame.pack(fill=tk.BOTH, expand=True)

        # Encabezado
        lbl_head = ttk.Label(
            pad_frame,
            text="Asistente de Mapeo de Columnas",
            font=("Segoe UI", 12, "bold"),
            foreground="#0f172a",
        )
        lbl_head.pack(anchor="w")

        lbl_desc = ttk.Label(
            pad_frame,
            text="Selecciona qué columna de tu Excel corresponde a cada campo de SAP Business One.\n"
                 "Puedes desmarcar cualquier campo si en esta ejecución no deseas modificarlo.",
            font=("Segoe UI", 9),
            foreground="#475569",
            justify=tk.LEFT,
        )
        lbl_desc.pack(anchor="w", pady=(2, 12))

        # Contenedor de filas de mapeo
        cards_frame = ttk.LabelFrame(pad_frame, text=" Campos de Planificación en SAP B1 ", padding="12")
        cards_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        # 1. Campo Código de Artículo
        f1 = ttk.Frame(cards_frame)
        f1.pack(fill=tk.X, pady=6)
        ttk.Label(f1, text="1. Código de Artículo (ItemCode):", font=("Segoe UI", 9, "bold"), width=32).pack(side=tk.LEFT)
        cb_item = ttk.Combobox(f1, textvariable=self.var_item, values=list(self.df.columns), state="readonly", width=26)
        cb_item.pack(side=tk.LEFT, padx=8)
        cb_item.bind("<<ComboboxSelected>>", lambda e: self._update_sample(self.var_item, self.sample_item))
        ttk.Label(f1, textvariable=self.sample_item, foreground="#0369a1", font=("Segoe UI", 9, "italic")).pack(side=tk.LEFT, padx=10)

        # 2. Cantidad de Pedido Mínimo
        f2 = ttk.Frame(cards_frame)
        f2.pack(fill=tk.X, pady=6)
        chk_min = ttk.Checkbutton(f2, text="2. Mínimo de Compra (Cantidad Mínima):", variable=self.flag_min, width=32, command=self._toggle_min)
        chk_min.pack(side=tk.LEFT)
        self.cb_min = ttk.Combobox(f2, textvariable=self.var_min, values=self.columns_list, state="readonly", width=26)
        self.cb_min.pack(side=tk.LEFT, padx=8)
        self.cb_min.bind("<<ComboboxSelected>>", lambda e: self._update_sample(self.var_min, self.sample_min))
        ttk.Label(f2, textvariable=self.sample_min, foreground="#0369a1", font=("Segoe UI", 9, "italic")).pack(side=tk.LEFT, padx=10)

        # 3. Tiempo Lead
        f3 = ttk.Frame(cards_frame)
        f3.pack(fill=tk.X, pady=6)
        chk_lead = ttk.Checkbutton(f3, text="3. Tiempo Lead (Días de Entrega):", variable=self.flag_lead, width=32, command=self._toggle_lead)
        chk_lead.pack(side=tk.LEFT)
        self.cb_lead = ttk.Combobox(f3, textvariable=self.var_lead, values=self.columns_list, state="readonly", width=26)
        self.cb_lead.pack(side=tk.LEFT, padx=8)
        self.cb_lead.bind("<<ComboboxSelected>>", lambda e: self._update_sample(self.var_lead, self.sample_lead))
        ttk.Label(f3, textvariable=self.sample_lead, foreground="#0369a1", font=("Segoe UI", 9, "italic")).pack(side=tk.LEFT, padx=10)

        # 4. Días de Tolerancia
        f4 = ttk.Frame(cards_frame)
        f4.pack(fill=tk.X, pady=6)
        chk_tol = ttk.Checkbutton(f4, text="4. Días de Tolerancia (Tiempo Retraso):", variable=self.flag_tol, width=32, command=self._toggle_tol)
        chk_tol.pack(side=tk.LEFT)
        self.cb_tol = ttk.Combobox(f4, textvariable=self.var_tol, values=self.columns_list, state="readonly", width=26)
        self.cb_tol.pack(side=tk.LEFT, padx=8)
        self.cb_tol.bind("<<ComboboxSelected>>", lambda e: self._update_sample(self.var_tol, self.sample_tol))
        ttk.Label(f4, textvariable=self.sample_tol, foreground="#0369a1", font=("Segoe UI", 9, "italic")).pack(side=tk.LEFT, padx=10)

        # Botones de Acción
        btn_box = ttk.Frame(pad_frame)
        btn_box.pack(fill=tk.X)

        btn_cancel = ttk.Button(btn_box, text="Cancelar", command=self.destroy)
        btn_cancel.pack(side=tk.RIGHT, padx=(10, 0))

        btn_confirm = ttk.Button(
            btn_box,
            text="✅ Confirmar Mapeo y Aplicar",
            style="ActionStart.TButton",
            command=self._apply_mapping,
        )
        btn_confirm.pack(side=tk.RIGHT)

    def _toggle_min(self):
        if not self.flag_min.get():
            self.cb_min.configure(state="disabled")
            self.sample_min.set("(Campo desactivado)")
        else:
            self.cb_min.configure(state="readonly")
            self._update_sample(self.var_min, self.sample_min)

    def _toggle_lead(self):
        if not self.flag_lead.get():
            self.cb_lead.configure(state="disabled")
            self.sample_lead.set("(Campo desactivado)")
        else:
            self.cb_lead.configure(state="readonly")
            self._update_sample(self.var_lead, self.sample_lead)

    def _toggle_tol(self):
        if not self.flag_tol.get():
            self.cb_tol.configure(state="disabled")
            self.sample_tol.set("(Campo desactivado)")
        else:
            self.cb_tol.configure(state="readonly")
            self._update_sample(self.var_tol, self.sample_tol)

    def _update_sample(self, var_col: tk.StringVar, sample_var: tk.StringVar):
        col = var_col.get()
        if not col or col == "(Omitir / No actualizar)" or col not in self.df.columns:
            sample_var.set("(Sin asignar)")
            return
        sample_val = ""
        for val in self.df[col]:
            if pd.notna(val) and str(val).strip():
                sample_val = str(val).strip()
                break
        if sample_val.endswith(".0"):
            sample_val = sample_val[:-2]
        sample_var.set(f'Ej: "{sample_val}"')

    def _update_all_samples(self):
        self._update_sample(self.var_item, self.sample_item)
        if self.flag_min.get():
            self.cb_min.configure(state="readonly")
            self._update_sample(self.var_min, self.sample_min)
        else:
            self.cb_min.configure(state="disabled")
            self.sample_min.set("(Campo desactivado)")

        if self.flag_lead.get():
            self.cb_lead.configure(state="readonly")
            self._update_sample(self.var_lead, self.sample_lead)
        else:
            self.cb_lead.configure(state="disabled")
            self.sample_lead.set("(Campo desactivado)")

        if self.flag_tol.get():
            self.cb_tol.configure(state="readonly")
            self._update_sample(self.var_tol, self.sample_tol)
        else:
            self.cb_tol.configure(state="disabled")
            self.sample_tol.set("(Campo desactivado)")

    def _apply_mapping(self):
        item_col = self.var_item.get()
        if not item_col or item_col not in self.df.columns:
            messagebox.showwarning("Atención", "Debe seleccionar una columna válida para el Código de Artículo.")
            return

        min_col = self.var_min.get() if self.flag_min.get() and self.var_min.get() != "(Omitir / No actualizar)" else ""
        lead_col = self.var_lead.get() if self.flag_lead.get() and self.var_lead.get() != "(Omitir / No actualizar)" else ""
        tol_col = self.var_tol.get() if self.flag_tol.get() and self.var_tol.get() != "(Omitir / No actualizar)" else ""

        mapping = {
            "item_col": item_col,
            "min_qty_col": min_col,
            "lead_time_col": lead_col,
            "tolerance_col": tol_col,
        }
        flags = {
            "min_qty": bool(min_col),
            "lead_time": bool(lead_col),
            "tolerance": bool(tol_col),
        }

        self.on_apply(mapping, flags)
        self.destroy()


class BotArticulosApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Bot de Artículos SAP B1 | Lead Time, Mínimo y Tolerancia")
        self.root.geometry("1000x880")
        self.root.minsize(940, 780)

        self._setup_styles()

        self.engine = BotArticulosEngine()
        self.engine.on_pause_callback = self._on_engine_pause_change
        self.root.bind_all("<Escape>", lambda e: self.toggle_pause())
        self.root.bind_all("<F8>", lambda e: self.toggle_pause())
        self.worker_thread = None
        self.calibrating = False

        # Variables de estado
        self.excel_path_var = tk.StringVar(value="")
        self.selected_window_var = tk.StringVar(value="")
        self.window_status_var = tk.StringVar(value="Verificando ventanas...")

        # Resumen de mapeo visible en la ventana principal
        self.lbl_mapped_item = tk.StringVar(value="Sin asignar")
        self.lbl_mapped_min = tk.StringVar(value="Sin asignar")
        self.lbl_mapped_lead = tk.StringVar(value="Sin asignar")
        self.lbl_mapped_tol = tk.StringVar(value="Sin asignar")

        # Variables de calibración visual (7 puntos)
        self.coord_btn_buscar_var = tk.StringVar(value="No calibrado")
        self.coord_item_code_var = tk.StringVar(value="No calibrado")
        self.coord_tab_var = tk.StringVar(value="No calibrado")
        self.coord_min_var = tk.StringVar(value="No calibrado")
        self.coord_lead_var = tk.StringVar(value="No calibrado")
        self.coord_tol_var = tk.StringVar(value="No calibrado")
        self.coord_btn_actualizar_var = tk.StringVar(value="No calibrado")

        self.speed_var = tk.StringVar(value="Normal (Recomendada)")
        self.status_bar_var = tk.StringVar(value="Listo. Cargue un archivo Excel para configurar el mapeo.")

        self._create_widgets()
        self._load_initial_coordinates()
        self.root.after(400, self.refresh_windows)

    def _setup_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("vista" if "vista" in style.theme_names() else "clam")
        except Exception:
            pass

        style.configure("Header.TLabel", font=("Segoe UI", 13, "bold"), foreground="#1e293b")
        style.configure("Subheader.TLabel", font=("Segoe UI", 9), foreground="#64748b")
        style.configure("StatusOk.TLabel", font=("Segoe UI", 9, "bold"), foreground="#15803d")
        style.configure("StatusErr.TLabel", font=("Segoe UI", 9, "bold"), foreground="#b91c1c")
        style.configure("CoordBadge.TLabel", font=("Segoe UI", 8, "bold"), foreground="#0369a1")
        style.configure("MapBadge.TLabel", font=("Segoe UI", 9, "bold"), foreground="#0f766e")
        style.configure("ActionStart.TButton", font=("Segoe UI", 10, "bold"))
        style.configure("ActionTest.TButton", font=("Segoe UI", 10, "bold"))
        style.configure("ActionStop.TButton", font=("Segoe UI", 10, "bold"))
        style.configure("MapButton.TButton", font=("Segoe UI", 9, "bold"))

    def _load_initial_coordinates(self):
        """Carga las coordenadas guardadas previamente en las variables visuales."""
        cfg = self.engine.config
        if cfg.get("btn_buscar_coord"):
            x, y = cfg["btn_buscar_coord"]
            self.coord_btn_buscar_var.set(f"X={x}, Y={y} ✓")
        if cfg.get("item_code_coord"):
            x, y = cfg["item_code_coord"]
            self.coord_item_code_var.set(f"X={x}, Y={y} ✓")
        if cfg.get("tab_planificacion_coord"):
            x, y = cfg["tab_planificacion_coord"]
            self.coord_tab_var.set(f"X={x}, Y={y} ✓")
        if cfg.get("min_qty_coord"):
            x, y = cfg["min_qty_coord"]
            self.coord_min_var.set(f"X={x}, Y={y} ✓")
        if cfg.get("lead_time_coord"):
            x, y = cfg["lead_time_coord"]
            self.coord_lead_var.set(f"X={x}, Y={y} ✓")
        if cfg.get("tolerance_coord"):
            x, y = cfg["tolerance_coord"]
            self.coord_tol_var.set(f"X={x}, Y={y} ✓")
        if cfg.get("btn_actualizar_coord"):
            x, y = cfg["btn_actualizar_coord"]
            self.coord_btn_actualizar_var.set(f"X={x}, Y={y} ✓")

    def _create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ------------------- ENCABEZADO -------------------
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 6))

        lbl_title = ttk.Label(
            header_frame,
            text="Actualizador de Planificación de Artículos (SAP B1 HANA)",
            style="Header.TLabel",
        )
        lbl_title.pack(anchor="w")

        lbl_sub = ttk.Label(
            header_frame,
            text="Automatización RPA para digitación masiva de Lead Time, Mínimo de Compra y Días de Tolerancia en 'Datos maestros de artículo'.",
            style="Subheader.TLabel",
        )
        lbl_sub.pack(anchor="w")

        ttk.Separator(main_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=(0, 6))

        # ------------------- TARJETA 1: ARCHIVO Y MAPEO DE COLUMNAS -------------------
        excel_card = ttk.LabelFrame(main_frame, text=" 1. Origen de Datos y Mapeo de Columnas ", padding="8")
        excel_card.pack(fill=tk.X, pady=(0, 6))

        file_row = ttk.Frame(excel_card)
        file_row.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(file_row, text="Archivo:", width=8).pack(side=tk.LEFT)
        ent_file = ttk.Entry(file_row, textvariable=self.excel_path_var, state="readonly")
        ent_file.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        btn_browse = ttk.Button(file_row, text="📁 Examinar Excel...", command=self.browse_file)
        btn_browse.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_open_map = ttk.Button(
            file_row,
            text="🔀 Mapear Columnas...",
            style="MapButton.TButton",
            command=self.open_mapping_dialog,
            state=tk.DISABLED,
        )
        self.btn_open_map.pack(side=tk.RIGHT)

        # Fila de Resumen de Mapeo
        map_summary_frame = ttk.Frame(excel_card)
        map_summary_frame.pack(fill=tk.X, pady=(2, 0))

        ttk.Label(map_summary_frame, text="Mapeo actual:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 8))

        ttk.Label(map_summary_frame, text="Art:").pack(side=tk.LEFT)
        ttk.Label(map_summary_frame, textvariable=self.lbl_mapped_item, style="MapBadge.TLabel").pack(side=tk.LEFT, padx=(2, 12))

        ttk.Label(map_summary_frame, text="Mínimo:").pack(side=tk.LEFT)
        ttk.Label(map_summary_frame, textvariable=self.lbl_mapped_min, style="MapBadge.TLabel").pack(side=tk.LEFT, padx=(2, 12))

        ttk.Label(map_summary_frame, text="Lead Time:").pack(side=tk.LEFT)
        ttk.Label(map_summary_frame, textvariable=self.lbl_mapped_lead, style="MapBadge.TLabel").pack(side=tk.LEFT, padx=(2, 12))

        ttk.Label(map_summary_frame, text="Tolerancia:").pack(side=tk.LEFT)
        ttk.Label(map_summary_frame, textvariable=self.lbl_mapped_tol, style="MapBadge.TLabel").pack(side=tk.LEFT, padx=(2, 12))

        self.lbl_row_count = ttk.Label(map_summary_frame, text="0 registros", foreground="#64748b")
        self.lbl_row_count.pack(side=tk.RIGHT)

        # ------------------- TARJETA 2: VENTANA SAP / RDP -------------------
        win_card = ttk.LabelFrame(main_frame, text=" 2. Ventana de SAP / Escritorio Remoto ", padding="6")
        win_card.pack(fill=tk.X, pady=(0, 6))

        win_row = ttk.Frame(win_card)
        win_row.pack(fill=tk.X)

        ttk.Label(win_row, text="Ventana:").pack(side=tk.LEFT, padx=(0, 5))
        self.cbo_windows = ttk.Combobox(
            win_row,
            textvariable=self.selected_window_var,
            state="readonly",
            width=58,
        )
        self.cbo_windows.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        btn_refresh = ttk.Button(win_row, text="🔄 Escanear Ventanas", command=self.refresh_windows)
        btn_refresh.pack(side=tk.RIGHT)

        status_row = ttk.Frame(win_card)
        status_row.pack(fill=tk.X, pady=(2, 0))
        self.lbl_win_status = ttk.Label(status_row, textvariable=self.window_status_var, style="StatusOk.TLabel")
        self.lbl_win_status.pack(side=tk.LEFT)

        # ------------------- TARJETA 3: CALIBRACIÓN ASISTIDA DE 7 PUNTOS -------------------
        calib_card = ttk.LabelFrame(main_frame, text=" 3. Calibración de Pantalla en SAP (7 Puntos) ", padding="6")
        calib_card.pack(fill=tk.X, pady=(0, 6))

        # Fila 1: Búsqueda y Código
        r1 = ttk.Frame(calib_card)
        r1.pack(fill=tk.X, pady=2)
        btn_b = ttk.Button(r1, text="🔍 1. Botón 'Buscar' (Lupa SAP)", width=32, command=lambda: self.start_calibration("btn_buscar"))
        btn_b.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(r1, textvariable=self.coord_btn_buscar_var, style="CoordBadge.TLabel", width=18).pack(side=tk.LEFT, padx=(0, 20))

        btn_c = ttk.Button(r1, text="📝 2. Campo 'Número de artículo'", width=32, command=lambda: self.start_calibration("item_code"))
        btn_c.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(r1, textvariable=self.coord_item_code_var, style="CoordBadge.TLabel", width=18).pack(side=tk.LEFT)

        # Fila 2: Pestaña y Mínimo
        r2 = ttk.Frame(calib_card)
        r2.pack(fill=tk.X, pady=2)
        btn_tab = ttk.Button(r2, text="📑 3. Pestaña 'Planificación'", width=32, command=lambda: self.start_calibration("tab"))
        btn_tab.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(r2, textvariable=self.coord_tab_var, style="CoordBadge.TLabel", width=18).pack(side=tk.LEFT, padx=(0, 20))

        btn_min = ttk.Button(r2, text="📦 4. Campo 'Pedido Mínimo'", width=32, command=lambda: self.start_calibration("min_qty"))
        btn_min.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(r2, textvariable=self.coord_min_var, style="CoordBadge.TLabel", width=18).pack(side=tk.LEFT)

        # Fila 3: Lead Time y Tolerancia
        r3 = ttk.Frame(calib_card)
        r3.pack(fill=tk.X, pady=2)
        btn_lead = ttk.Button(r3, text="⏱️ 5. Campo 'Tiempo Lead'", width=32, command=lambda: self.start_calibration("lead"))
        btn_lead.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(r3, textvariable=self.coord_lead_var, style="CoordBadge.TLabel", width=18).pack(side=tk.LEFT, padx=(0, 20))

        btn_tol = ttk.Button(r3, text="⏳ 6. Campo 'Días Tolerancia'", width=32, command=lambda: self.start_calibration("tol"))
        btn_tol.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(r3, textvariable=self.coord_tol_var, style="CoordBadge.TLabel", width=18).pack(side=tk.LEFT)

        # Fila 4: Guardar y Acciones Rápidas
        r4 = ttk.Frame(calib_card)
        r4.pack(fill=tk.X, pady=2)
        btn_act = ttk.Button(r4, text="💾 7. Botón 'Buscar / Actualizar'", width=32, command=lambda: self.start_calibration("btn_actualizar"))
        btn_act.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(r4, textvariable=self.coord_btn_actualizar_var, style="CoordBadge.TLabel", width=18).pack(side=tk.LEFT, padx=(0, 20))

        btn_save_coords = ttk.Button(r4, text="💾 Guardar Coordenadas", command=self._manual_save_coords)
        btn_save_coords.pack(side=tk.LEFT, padx=(0, 8))

        # Fila 5: Asistente Secuencial y Velocidad
        r5 = ttk.Frame(calib_card)
        r5.pack(fill=tk.X, pady=(4, 0))
        btn_wizard = ttk.Button(
            r5,
            text="🚀 Asistente Rápido: Calibrar los 7 puntos en secuencia (35 seg)",
            command=self.start_wizard_calibration,
        )
        btn_wizard.pack(side=tk.LEFT, padx=(0, 20))

        ttk.Label(r5, text="Velocidad RDP:").pack(side=tk.LEFT, padx=(0, 4))
        cbo_speed = ttk.Combobox(
            r5,
            textvariable=self.speed_var,
            values=["Rápida", "Normal (Recomendada)", "Lenta (Conexión inestable)"],
            state="readonly",
            width=22,
        )
        cbo_speed.pack(side=tk.LEFT)

        # ------------------- TARJETA 4: PROGRESO Y CONTROLES -------------------
        exec_card = ttk.LabelFrame(main_frame, text=" 4. Ejecución y Resultados ", padding="8")
        exec_card.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        btn_bar = ttk.Frame(exec_card)
        btn_bar.pack(fill=tk.X, pady=(0, 6))

        # Botón de Prueba Individual (NUEVO)
        self.btn_test = ttk.Button(
            btn_bar,
            text="🧪 PROBAR CON 1 ARTÍCULO",
            command=self.start_single_test,
            style="ActionTest.TButton",
            width=28,
        )
        self.btn_test.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_start = ttk.Button(
            btn_bar,
            text="▶ INICIAR ACTUALIZACIÓN MASIVA",
            command=self.start_process,
            style="ActionStart.TButton",
            width=34,
        )
        self.btn_start.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_pause = ttk.Button(
            btn_bar,
            text="⏸ Pausar (ESC)",
            command=self.toggle_pause,
            state=tk.DISABLED,
            width=16,
        )
        self.btn_pause.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_stop = ttk.Button(
            btn_bar,
            text="⏹ Detener",
            command=self.stop_process,
            state=tk.DISABLED,
            style="ActionStop.TButton",
            width=14,
        )
        self.btn_stop.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_export = ttk.Button(
            btn_bar,
            text="💾 Exportar Resultados",
            command=self.export_results,
            state=tk.DISABLED,
        )
        self.btn_export.pack(side=tk.RIGHT)

        prog_frame = ttk.Frame(exec_card)
        prog_frame.pack(fill=tk.X, pady=(0, 4))

        self.prog_bar = ttk.Progressbar(prog_frame, orient=tk.HORIZONTAL, mode="determinate")
        self.prog_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        self.lbl_progress_pct = ttk.Label(prog_frame, text="0 / 0 (0%)", width=14)
        self.lbl_progress_pct.pack(side=tk.RIGHT)

        tree_frame = ttk.Frame(exec_card)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        columns = ("num", "item_code", "min_qty", "lead_time", "tolerance", "status", "message")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=7)

        self.tree.heading("num", text="#")
        self.tree.heading("item_code", text="Artículo")
        self.tree.heading("min_qty", text="Pedido Mínimo")
        self.tree.heading("lead_time", text="Tiempo Lead")
        self.tree.heading("tolerance", text="Tolerancia (Días)")
        self.tree.heading("status", text="Estado")
        self.tree.heading("message", text="Detalle")

        self.tree.column("num", width=35, anchor=tk.CENTER)
        self.tree.column("item_code", width=120, anchor=tk.CENTER)
        self.tree.column("min_qty", width=100, anchor=tk.CENTER)
        self.tree.column("lead_time", width=90, anchor=tk.CENTER)
        self.tree.column("tolerance", width=110, anchor=tk.CENTER)
        self.tree.column("status", width=85, anchor=tk.CENTER)
        self.tree.column("message", width=340, anchor=tk.W)

        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.tag_configure("OK", foreground="#15803d")
        self.tree.tag_configure("ERROR", foreground="#b91c1c")
        self.tree.tag_configure("PROCESANDO", foreground="#2563eb")

        status_bar = ttk.Frame(main_frame)
        status_bar.pack(fill=tk.X, pady=(2, 0))
        lbl_bot_status = ttk.Label(status_bar, textvariable=self.status_bar_var, relief=tk.SUNKEN, anchor="w", padding="3")
        lbl_bot_status.pack(fill=tk.X)

    # ------------------- MÉTODOS DE MAPEO Y ARCHIVO -------------------

    def refresh_windows(self):
        windows = window_manager.list_visible_windows()
        window_titles = [t for _, t in windows]
        self.cbo_windows["values"] = window_titles

        target = window_manager.find_target_window()
        if target:
            _, title = target
            self.selected_window_var.set(title)
            self.window_status_var.set(f"🟢 Detectado: {title}")
            self.lbl_win_status.configure(style="StatusOk.TLabel")
        else:
            self.window_status_var.set("🔴 No se detectó SAP ni Escritorio Remoto abierto.")
            self.lbl_win_status.configure(style="StatusErr.TLabel")

    def browse_file(self):
        filepath = filedialog.askopenfilename(
            title="Seleccionar archivo de artículos",
            filetypes=[
                ("Archivos de datos", "*.xlsx *.xls *.csv"),
                ("Excel (.xlsx)", "*.xlsx"),
                ("Excel 97-2003 (.xls)", "*.xls"),
                ("Archivos CSV", "*.csv"),
                ("Todos los archivos", "*.*"),
            ],
        )
        if not filepath:
            return

        try:
            df, detected = self.engine.load_file(filepath)
            self.excel_path_var.set(filepath)
            self.btn_open_map.config(state=tk.NORMAL)

            total_rows = len(df)
            self.lbl_row_count.config(text=f"✓ {total_rows} artículos", foreground="#15803d")
            self.status_bar_var.set(f"Archivo cargado: {os.path.basename(filepath)}. Abriendo asistente de mapeo...")

            # Abrir automáticamente el diálogo de mapeo para confirmar columnas
            self.open_mapping_dialog()

        except Exception as e:
            messagebox.showerror("Error al cargar archivo", f"No se pudo leer el archivo:\n{e}")

    def open_mapping_dialog(self):
        if self.engine.dataframe is None:
            messagebox.showwarning("Atención", "Cargue primero un archivo Excel.")
            return

        current_mapping = {
            "item_col": self.engine.item_col,
            "min_qty_col": self.engine.min_qty_col or "(Omitir / No actualizar)",
            "lead_time_col": self.engine.lead_time_col or "(Omitir / No actualizar)",
            "tolerance_col": self.engine.tolerance_col or "(Omitir / No actualizar)",
        }
        current_flags = self.engine.update_flags

        ColumnMappingDialog(
            parent=self.root,
            df=self.engine.dataframe,
            current_mapping=current_mapping,
            current_flags=current_flags,
            on_apply_callback=self._apply_mapped_columns,
        )

    def _apply_mapped_columns(self, mapping: dict, flags: dict):
        self.engine.item_col = mapping["item_col"]
        self.engine.min_qty_col = mapping["min_qty_col"]
        self.engine.lead_time_col = mapping["lead_time_col"]
        self.engine.tolerance_col = mapping["tolerance_col"]
        self.engine.update_flags = flags

        # Actualizar etiquetas de resumen en la interfaz
        self.lbl_mapped_item.set(mapping["item_col"] or "Sin asignar")
        self.lbl_mapped_min.set(mapping["min_qty_col"] if flags.get("min_qty") else "(Omitido)")
        self.lbl_mapped_lead.set(mapping["lead_time_col"] if flags.get("lead_time") else "(Omitido)")
        self.lbl_mapped_tol.set(mapping["tolerance_col"] if flags.get("tolerance") else "(Omitido)")

        # Refrescar tabla de previsualización
        df = self.engine.dataframe
        self.tree.delete(*self.tree.get_children())

        for idx, row in df.iterrows():
            i_val = str(row[mapping["item_col"]]) if pd.notna(row[mapping["item_col"]]) else ""
            m_val = str(row[mapping["min_qty_col"]]) if flags.get("min_qty") and mapping["min_qty_col"] in df.columns and pd.notna(row[mapping["min_qty_col"]]) else "(Omitido)"
            l_val = str(row[mapping["lead_time_col"]]) if flags.get("lead_time") and mapping["lead_time_col"] in df.columns and pd.notna(row[mapping["lead_time_col"]]) else "(Omitido)"
            t_val = str(row[mapping["tolerance_col"]]) if flags.get("tolerance") and mapping["tolerance_col"] in df.columns and pd.notna(row[mapping["tolerance_col"]]) else "(Omitido)"

            if m_val.endswith(".0"): m_val = m_val[:-2]
            if l_val.endswith(".0"): l_val = l_val[:-2]
            if t_val.endswith(".0"): t_val = t_val[:-2]

            self.tree.insert("", tk.END, iid=f"row_{idx}", values=(idx + 1, i_val, m_val, l_val, t_val, "Listo", "En espera"))

        total_rows = len(df)
        self.prog_bar["maximum"] = total_rows
        self.prog_bar["value"] = 0
        self.lbl_progress_pct.config(text=f"0 / {total_rows} (0%)")
        self.status_bar_var.set("Mapeo aplicado exitosamente. Verifique la calibración antes de iniciar.")

    # ------------------- CALIBRACIÓN DE 7 PUNTOS -------------------

    def start_calibration(self, target_point: str, on_complete: Optional[callable] = None):
        if self.calibrating:
            return
        self.calibrating = True

        names = {
            "btn_buscar": ("1. Botón 'Buscar' (Lupa barra de herramientas SAP)", self.coord_btn_buscar_var, "btn_buscar_coord"),
            "item_code": ("2. Campo 'Número de artículo' (Caja de texto de código)", self.coord_item_code_var, "item_code_coord"),
            "tab": ("3. Pestaña 'Datos de Planificación'", self.coord_tab_var, "tab_planificacion_coord"),
            "min_qty": ("4. Campo 'Cantidad de Pedido Mínimo'", self.coord_min_var, "min_qty_coord"),
            "lead": ("5. Campo 'Tiempo Lead'", self.coord_lead_var, "lead_time_coord"),
            "tol": ("6. Campo 'Días de Tolerancia'", self.coord_tol_var, "tolerance_coord"),
            "btn_actualizar": ("7. Botón 'Buscar / Actualizar' (inferior izquierdo)", self.coord_btn_actualizar_var, "btn_actualizar_coord"),
        }

        label_name, text_var, config_key = names[target_point]

        calib_win = tk.Toplevel(self.root)
        calib_win.title(f"Calibrador: {label_name}")
        calib_win.geometry("480x170")
        calib_win.attributes("-topmost", True)
        calib_win.resizable(False, False)

        lbl_info = ttk.Label(
            calib_win,
            text=f"Coloca el puntero del mouse exactamente sobre:\n'{label_name}' en tu SAP...",
            font=("Segoe UI", 10),
            justify=tk.CENTER,
            padding=10,
        )
        lbl_info.pack()

        lbl_countdown = ttk.Label(calib_win, text="5", font=("Segoe UI", 28, "bold"), foreground="#e11d48")
        lbl_countdown.pack()

        def countdown_step(remaining):
            if remaining > 0:
                lbl_countdown.config(text=str(remaining))
                calib_win.after(1000, countdown_step, remaining - 1)
            else:
                window_manager.ensure_desktop_attached()
                import win32api
                x, y = win32api.GetCursorPos()
                calib_win.destroy()
                self.calibrating = False

                self.engine.config[config_key] = (x, y)
                text_var.set(f"X={x}, Y={y} ✓")
                self.engine.save_coordinates_to_file()
                self.root.bell()

                if on_complete:
                    on_complete()
                else:
                    messagebox.showinfo("Calibración Exitosa", f"Se guardó la posición de:\n'{label_name}'\nX={x}, Y={y}")

        calib_win.after(1000, countdown_step, 4)

    def start_wizard_calibration(self):
        steps = ["btn_buscar", "item_code", "tab", "min_qty", "lead", "tol", "btn_actualizar"]
        step_names = [
            "1. Botón 'Buscar' (Lupa barra SAP)",
            "2. Campo 'Número de artículo' (Código)",
            "3. Pestaña 'Datos de Planificación'",
            "4. Campo 'Cantidad Pedido Mínimo'",
            "5. Campo 'Tiempo Lead'",
            "6. Campo 'Días de Tolerancia'",
            "7. Botón 'Buscar / Actualizar' (inferior izquierdo)",
        ]

        def run_step(step_idx):
            if step_idx < len(steps):
                self.start_calibration(steps[step_idx], on_complete=lambda: run_step(step_idx + 1))
            else:
                messagebox.showinfo(
                    "Asistente Completo",
                    "¡Los 7 puntos fueron calibrados con éxito!\n\n"
                    "Las coordenadas se guardaron en disco y estarán disponibles en todas tus sesiones."
                )

        confirm = messagebox.askyesno(
            "Asistente de Calibración de 7 Puntos",
            "El asistente te pedirá ubicar el puntero del mouse en 7 puntos sucesivos (5 segundos para cada uno):\n\n"
            "1. 🔍 Botón 'Buscar' (Lupa en barra de SAP)\n"
            "2. 📝 Campo 'Número de artículo'\n"
            "3. 📑 Pestaña 'Datos de Planificación'\n"
            "4. 📦 Campo 'Cantidad Pedido Mínimo'\n"
            "5. ⏱️ Campo 'Tiempo Lead'\n"
            "6. ⏳ Campo 'Días de Tolerancia'\n"
            "7. 💾 Botón 'Buscar / Actualizar' (inferior izquierdo)\n\n"
            "¿Deseas iniciar?",
        )
        if confirm:
            run_step(0)

    def _manual_save_coords(self):
        saved = self.engine.save_coordinates_to_file()
        if saved:
            messagebox.showinfo("Guardado", "Coordenadas guardadas permanentemente en disco.")
        else:
            messagebox.showerror("Error", "No se pudieron guardar las coordenadas en disco.")

    def _apply_settings(self):
        speed = self.speed_var.get()
        if "Rápida" in speed:
            self.engine.config["delay_after_find"] = 0.35
            self.engine.config["delay_after_search"] = 0.75
            self.engine.config["delay_after_tab_click"] = 0.3
            self.engine.config["delay_between_fields"] = 0.18
            self.engine.config["delay_after_save"] = 0.65
        elif "Lenta" in speed:
            self.engine.config["delay_after_find"] = 0.65
            self.engine.config["delay_after_search"] = 1.4
            self.engine.config["delay_after_tab_click"] = 0.55
            self.engine.config["delay_between_fields"] = 0.35
            self.engine.config["delay_after_save"] = 1.2
        else:  # Normal
            self.engine.config["delay_after_find"] = 0.45
            self.engine.config["delay_after_search"] = 1.0
            self.engine.config["delay_after_tab_click"] = 0.4
            self.engine.config["delay_between_fields"] = 0.25
            self.engine.config["delay_after_save"] = 0.8

    # ------------------- MODO DE PRUEBA: 1 ARTÍCULO -------------------

    def start_single_test(self):
        if not self.excel_path_var.get() or self.engine.dataframe is None:
            messagebox.showwarning("Atención", "Por favor seleccione un archivo Excel antes de probar.")
            return

        cfg = self.engine.config
        if not cfg.get("item_code_coord") or not cfg.get("tab_planificacion_coord"):
            messagebox.showwarning(
                "Calibración Incompleta",
                "Para la prueba se requiere al menos calibrar:\n"
                "1. Campo 'Número de artículo'\n"
                "2. Pestaña 'Datos de Planificación'\n\n"
                "Usa el botón de calibración o el Asistente Rápido."
            )
            return

        custom_title = self.selected_window_var.get().strip() or None
        target = window_manager.find_target_window(custom_title)
        if not target:
            messagebox.showerror(
                "SAP no detectado",
                "No se encontró la ventana de SAP Business One ni Escritorio Remoto abierto."
            )
            return

        self.engine.config["custom_window_title"] = custom_title
        self._apply_settings()

        self.btn_test.config(state=tk.DISABLED)
        self.btn_start.config(state=tk.DISABLED)
        self.btn_pause.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)

        def worker():
            try:
                result = self.engine.run_single_test_item(
                    item_index=0,
                    on_countdown_tick=self._on_countdown,
                )
                self.root.after(0, lambda: self._on_single_test_complete(result))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Error en Prueba", f"Ocurrió un error:\n{e}"))
                self.root.after(0, self._reset_controls)

        self.worker_thread = threading.Thread(target=worker, daemon=True)
        self.worker_thread.start()

    def _on_single_test_complete(self, result: dict):
        self._reset_controls()
        item_code = result.get("item_code", "")
        actualizados = ", ".join(result.get("actualizados", [])) or "Sin campos modificados"

        self.status_bar_var.set(f"Prueba completada para {item_code}. Revisa la pantalla de SAP.")

        resp = messagebox.askyesno(
            "Verificación de Prueba de 1 Artículo",
            f"✅ Se ejecutó la prueba con el artículo: {item_code}\n\n"
            f"Valores digitados: {actualizados}\n\n"
            f"Por favor verifica en tu pantalla de SAP:\n"
            f"1. ¿Hizo clic en Modo Buscar y luego en 'Número de artículo'?\n"
            f"2. ¿Pegó el código en la casilla correcta sin tocar otros campos?\n"
            f"3. ¿Entró a la pestaña 'Datos de planificación'?\n"
            f"4. ¿Digitó los valores en sus casillas correspondientes?\n"
            f"5. ¿Se guardaron los cambios?\n\n"
            f"¿Todo se ejecutó correctamente y deseas INICIAR la actualización masiva ahora?",
        )
        if resp:
            self.start_process()

    # ------------------- PROCESO MASIVO COMPLETO -------------------

    def start_process(self):
        if not self.excel_path_var.get() or self.engine.dataframe is None:
            messagebox.showwarning("Atención", "Por favor seleccione un archivo Excel antes de iniciar.")
            return

        cfg = self.engine.config
        if not cfg.get("item_code_coord") or not cfg.get("tab_planificacion_coord"):
            messagebox.showwarning(
                "Calibración Requerida",
                "Es necesario calibrar al menos el campo 'Número de artículo' y la pestaña 'Datos de Planificación'.\n\n"
                "Usa el Asistente Rápido para calibrar los puntos en secuencia.",
            )
            return

        custom_title = self.selected_window_var.get().strip() or None
        target = window_manager.find_target_window(custom_title)
        if not target:
            messagebox.showerror(
                "Error: SAP / Escritorio Remoto no detectado",
                "⚠️ NO SE ENCONTRÓ LA VENTANA DE SAP BUSINESS ONE NI DE ESCRITORIO REMOTO.\n\n"
                "Para poder ejecutar la actualización:\n"
                "1. Abre la conexión de Escritorio Remoto (182.160.29.90).\n"
                "2. Inicia sesión en SAP Business One con tu usuario.\n"
                "3. Abre la pantalla 'Datos maestros de artículo'.\n"
                "4. Vuelve a hacer clic en Iniciar.",
            )
            return

        self.engine.config["custom_window_title"] = custom_title
        self._apply_settings()

        self.btn_test.config(state=tk.DISABLED)
        self.btn_start.config(state=tk.DISABLED)
        self.btn_pause.config(state=tk.NORMAL, text="⏸ Pausar (ESC)", style="TButton")
        self.btn_stop.config(state=tk.NORMAL)
        self.btn_export.config(state=tk.DISABLED)

        self.worker_thread = threading.Thread(target=self._run_worker, daemon=True)
        self.worker_thread.start()

    def _run_worker(self):
        try:
            self.engine.run_update_items(
                progress_callback=self._on_row_progress,
                on_countdown_tick=self._on_countdown,
                completion_callback=self._on_sync_finished,
            )
        except SAPWindowNotFoundError as e:
            self.root.after(0, lambda: messagebox.showerror("Error de Ventana SAP", str(e)))
            self.root.after(0, self._reset_controls)
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error en el Proceso", f"Ocurrió un error inesperado:\n{e}"))
            self.root.after(0, self._reset_controls)

    def _on_countdown(self, seconds_left: int):
        if seconds_left > 0:
            self.root.after(
                0,
                lambda s=seconds_left: self.status_bar_var.set(
                    f"⏳ Enfocando SAP... Iniciando en {s} segundos..."
                ),
            )
        else:
            self.root.after(0, lambda: self.status_bar_var.set("▶ Actualización de artículos en curso..."))

    def _on_row_progress(self, data: dict):
        def update_ui():
            idx = data["index"]
            total = data["total"]
            row_idx = data["row_idx"]
            status = data["status"]
            msg = data["message"]

            pct = int((idx / total) * 100) if total > 0 else 0
            self.prog_bar["value"] = idx
            self.lbl_progress_pct.config(text=f"{idx} / {total} ({pct}%)")
            self.status_bar_var.set(f"[{idx}/{total}] Artículo {data['item_code']} -> {msg}")

            item_id = f"row_{row_idx}"
            if self.tree.exists(item_id):
                self.tree.item(
                    item_id,
                    values=(idx, data["item_code"], data["min_val"], data["lead_val"], data["tol_val"], status, msg),
                    tags=(status,),
                )
                self.tree.see(item_id)

        self.root.after(0, update_ui)

    def _on_sync_finished(self, stats: dict):
        def finalize():
            self._reset_controls()
            self.btn_export.config(state=tk.NORMAL)

            if stats.get("aborted"):
                self.status_bar_var.set("⏹ Proceso detenido por el usuario.")
                messagebox.showinfo("Proceso Detenido", f"La actualización fue detenida.\n\nArtículos procesados: {stats['success']}/{stats['total']}")
            else:
                self.status_bar_var.set(f"✓ Finalizado: {stats['success']} actualizados, {stats['failed']} errores.")
                messagebox.showinfo(
                    "Actualización Completada",
                    f"¡Proceso finalizado exitosamente!\n\n"
                    f"Total artículos: {stats['total']}\n"
                    f"Actualizados con éxito: {stats['success']}\n"
                    f"Errores: {stats['failed']}\n\n"
                    "Puedes hacer clic en 'Exportar Resultados' para guardar la bitácora en Excel.",
                )

        self.root.after(0, finalize)

    def _on_engine_pause_change(self, is_paused: bool):
        def _update():
            if is_paused:
                self.btn_pause.config(text="▶ Reanudar (ESC)", style="ActionStart.TButton")
                self.status_bar_var.set("⏸ PAUSADO POR USUARIO (Presiona ESC o clic en Reanudar)")
            else:
                self.btn_pause.config(text="⏸ Pausar (ESC)", style="TButton")
                self.status_bar_var.set("▶ Procesando artículos...")
        self.root.after(0, _update)

    def _reset_controls(self):
        self.btn_test.config(state=tk.NORMAL)
        self.btn_start.config(state=tk.NORMAL)
        self.btn_pause.config(state=tk.DISABLED, text="⏸ Pausar (ESC)", style="TButton")
        self.btn_stop.config(state=tk.DISABLED)

    def toggle_pause(self):
        if not self.engine.is_running:
            return
        if self.engine.is_paused():
            self.engine.resume()
            self._on_engine_pause_change(False)
        else:
            self.engine.pause()
            self._on_engine_pause_change(True)

    def stop_process(self):
        self.engine.stop()
        self.status_bar_var.set("Deteniendo bot...")

    def export_results(self):
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel (.xlsx)", "*.xlsx")],
            initialfile="resultado_actualizacion_articulos.xlsx",
            title="Guardar resultados de artículos",
        )
        if not save_path:
            return

        try:
            records = []
            for item in self.tree.get_children():
                vals = self.tree.item(item)["values"]
                records.append({
                    "#": vals[0],
                    "Codigo_Articulo": vals[1],
                    "Minimo_Compra": vals[2],
                    "Tiempo_Lead": vals[3],
                    "Dias_Tolerancia": vals[4],
                    "Estado": vals[5],
                    "Detalle": vals[6],
                })
            out_df = pd.DataFrame(records)
            out_df.to_excel(save_path, index=False)
            messagebox.showinfo("Exportación Exitosa", f"Archivo guardado exitosamente en:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Error al exportar", f"No se pudo guardar el archivo:\n{e}")


def main():
    root = tk.Tk()
    app = BotArticulosApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
