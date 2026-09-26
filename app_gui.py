"""
app_gui.py
Interfaz Gráfica (Tkinter) para el Bot de Actualización de Clientes / Socios de Negocios en SAP B1 HANA.
Incluye:
  - Ventana modal interactiva de Mapeo de Columnas Excel ➔ Campos de Clientes en SAP.
  - Soporte de campos: RTN, Teléfono 1, Teléfono móvil, Correo, Activo, WBCUSTID, SyncFlag,
    ID de dirección, Calle/Número, Ciudad e Indicador de impuestos.
  - Calibración por categorías y Asistente Rápido secuencial con persistencia en JSON.
  - Protocolo de auto-recuperación ante errores al Actualizar (Crear ➔ Descartar ➔ Buscar).
  - Botón de 'Probar con 1 cliente' antes de corridas masivas.
  - Ejecución continua segura con pausas (F8 / ESC) y registro en vivo.

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
from typing import Optional, Dict
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import win32api

import window_manager
from bot_engine import BotEngine, SAPWindowNotFoundError


class ClientColumnMappingDialog(tk.Toplevel):
    """
    Ventana emergente interactiva para mapear las columnas del Excel cargado
    con los campos de Socios de Negocios en SAP Business One.
    """
    def __init__(self, parent, df: pd.DataFrame, current_mapping: dict, current_flags: dict, on_apply_callback):
        super().__init__(parent)
        self.title("Mapeo de Columnas: Excel ➔ Campos de Clientes en SAP B1")
        self.geometry("780x680")
        self.minsize(740, 620)
        self.transient(parent)
        self.grab_set()

        self.df = df
        self.on_apply = on_apply_callback
        self.columns_list = ["(Omitir / No actualizar)"] + list(df.columns)

        # Variables de selección de columnas
        self.vars = {
            "card_code": tk.StringVar(value=current_mapping.get("card_code", "")),
            "rtn": tk.StringVar(value=current_mapping.get("rtn", "(Omitir / No actualizar)")),
            "telefono": tk.StringVar(value=current_mapping.get("telefono", "(Omitir / No actualizar)")),
            "movil": tk.StringVar(value=current_mapping.get("movil", "(Omitir / No actualizar)")),
            "correo": tk.StringVar(value=current_mapping.get("correo", "(Omitir / No actualizar)")),
            "activo": tk.StringVar(value=current_mapping.get("activo", "(Omitir / No actualizar)")),
            "wbcustid": tk.StringVar(value=current_mapping.get("wbcustid", "(Omitir / No actualizar)")),
            "syncflag": tk.StringVar(value=current_mapping.get("syncflag", "(Omitir / No actualizar)")),
            "id_direccion": tk.StringVar(value=current_mapping.get("id_direccion", "(Omitir / No actualizar)")),
            "calle_numero": tk.StringVar(value=current_mapping.get("calle_numero", "(Omitir / No actualizar)")),
            "ciudad": tk.StringVar(value=current_mapping.get("ciudad", "(Omitir / No actualizar)")),
            "indicador_impuestos": tk.StringVar(value=current_mapping.get("indicador_impuestos", "(Omitir / No actualizar)")),
        }

        # Variables de activación de campos
        self.flags = {
            k: tk.BooleanVar(value=current_flags.get(k, True)) for k in self.vars if k != "card_code"
        }

        # Variables para muestras en vivo
        self.samples = {k: tk.StringVar() for k in self.vars}

        # Controles Combobox guardados para toggle
        self.combos: Dict[str, ttk.Combobox] = {}

        self._build_ui()
        self._update_all_samples()

    def _build_ui(self):
        pad_frame = ttk.Frame(self, padding="15")
        pad_frame.pack(fill=tk.BOTH, expand=True)

        lbl_head = ttk.Label(
            pad_frame,
            text="Asistente de Mapeo de Clientes (SAP B1)",
            font=("Segoe UI", 12, "bold"),
            foreground="#0f172a",
        )
        lbl_head.pack(anchor="w")

        lbl_desc = ttk.Label(
            pad_frame,
            text="Selecciona qué columna de tu Excel corresponde a cada campo de SAP Business One.\n"
                 "Puedes desmarcar las casillas de los campos que no desees modificar en esta ejecución.",
            font=("Segoe UI", 9),
            foreground="#475569",
            justify=tk.LEFT,
        )
        lbl_desc.pack(anchor="w", pady=(2, 8))

        canvas = tk.Canvas(pad_frame, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(pad_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0, 10))
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, before=canvas)

        # ------------------- SECCIÓN 1: IDENTIFICADOR -------------------
        f_id = ttk.LabelFrame(scrollable_frame, text=" 1. Identificador Principal (Cabecera) ", padding="8")
        f_id.pack(fill=tk.X, expand=True, pady=(0, 6), padx=4)

        r_code = ttk.Frame(f_id)
        r_code.pack(fill=tk.X, pady=3)
        ttk.Label(r_code, text="Código de Cliente (CardCode):*", font=("Segoe UI", 9, "bold"), width=32).pack(side=tk.LEFT)
        cb_code = ttk.Combobox(r_code, textvariable=self.vars["card_code"], values=list(self.df.columns), state="readonly", width=24)
        cb_code.pack(side=tk.LEFT, padx=6)
        cb_code.bind("<<ComboboxSelected>>", lambda e: self._update_sample("card_code"))
        ttk.Label(r_code, textvariable=self.samples["card_code"], foreground="#0369a1", font=("Segoe UI", 9, "italic")).pack(side=tk.LEFT, padx=8)

        # ------------------- SECCIÓN 2: CABECERA Y GENERAL -------------------
        f_gen = ttk.LabelFrame(scrollable_frame, text=" 2. Cabecera, Pestaña General y UDF ", padding="8")
        f_gen.pack(fill=tk.X, expand=True, pady=(0, 6), padx=4)

        gen_fields = [
            ("rtn", "RTN (Número de Identificación Fiscal):"),
            ("telefono", "Teléfono 1 (Pestaña General):"),
            ("movil", "Teléfono Móvil (Pestaña General):"),
            ("correo", "Correo Electrónico (Pestaña General):"),
            ("activo", "Estado Activo (Pestaña General - Y/N):"),
            ("wbcustid", "WBCUSTID (ID Zoho - Panel UDF):"),
            ("syncflag", "SyncFlag (Bandera Sincronización UDF):"),
        ]

        for key, label in gen_fields:
            self._create_field_row(f_gen, key, label)

        # ------------------- SECCIÓN 3: PESTAÑA DIRECCIONES -------------------
        f_dir = ttk.LabelFrame(scrollable_frame, text=" 3. Pestaña Direcciones ", padding="8")
        f_dir.pack(fill=tk.X, expand=True, pady=(0, 6), padx=4)

        dir_fields = [
            ("id_direccion", "ID de Dirección (ej. San Pedro Sula):"),
            ("calle_numero", "Calle / Número (Línea de Dirección):"),
            ("ciudad", "Ciudad (Municipio / Localidad):"),
            ("indicador_impuestos", "Indicador de Impuestos (ej. ISV):"),
        ]

        for key, label in dir_fields:
            self._create_field_row(f_dir, key, label)

        # ------------------- BOTONES DE ACCIÓN -------------------
        btn_box = ttk.Frame(pad_frame)
        btn_box.pack(fill=tk.X, pady=(6, 0))

        btn_cancel = ttk.Button(btn_box, text="Cancelar", command=self.destroy)
        btn_cancel.pack(side=tk.RIGHT, padx=(10, 0))

        btn_confirm = ttk.Button(
            btn_box,
            text="✅ Confirmar Mapeo y Aplicar",
            style="ActionStart.TButton",
            command=self._apply_mapping,
        )
        btn_confirm.pack(side=tk.RIGHT)

    def _create_field_row(self, parent, key: str, label_text: str):
        row = ttk.Frame(parent)
        row.pack(fill=tk.X, pady=3)

        chk = ttk.Checkbutton(
            row,
            text=label_text,
            variable=self.flags[key],
            width=36,
            command=lambda k=key: self._toggle_field(k),
        )
        chk.pack(side=tk.LEFT)

        cb = ttk.Combobox(row, textvariable=self.vars[key], values=self.columns_list, state="readonly", width=24)
        cb.pack(side=tk.LEFT, padx=6)
        cb.bind("<<ComboboxSelected>>", lambda e, k=key: self._update_sample(k))
        self.combos[key] = cb

        ttk.Label(row, textvariable=self.samples[key], foreground="#0369a1", font=("Segoe UI", 9, "italic")).pack(side=tk.LEFT, padx=8)

    def _toggle_field(self, key: str):
        active = self.flags[key].get()
        cb = self.combos.get(key)
        if cb:
            if active:
                cb.configure(state="readonly")
                self._update_sample(key)
            else:
                cb.configure(state="disabled")
                self.samples[key].set("(Campo desactivado)")

    def _update_sample(self, key: str):
        col = self.vars[key].get()
        if not col or col == "(Omitir / No actualizar)" or col not in self.df.columns:
            self.samples[key].set("(Sin asignar)")
            return
        sample_val = ""
        for val in self.df[col]:
            if pd.notna(val) and str(val).strip():
                sample_val = str(val).strip()
                break
        if sample_val:
            if key == "activo":
                if sample_val.upper() in ["Y", "S", "SI", "1", "1.0", "TRUE", "T", "ACTIVO"]:
                    sample_val = "Y"
                elif sample_val.upper() in ["N", "NO", "0", "0.0", "FALSE", "F", "INACTIVO"]:
                    sample_val = "N"
            elif key in ["telefono", "movil", "wbcustid", "id_direccion", "rtn", "card_code"]:
                try:
                    c = sample_val.replace(",", ".")
                    if c.count(".") > 1:
                        parts = c.split(".")
                        c = "".join(parts[:-1]) + "." + parts[-1]
                    f = float(c)
                    sample_val = str(int(round(f)))
                except Exception:
                    if sample_val.endswith(".0"):
                        sample_val = sample_val[:-2]
            elif sample_val.endswith(".0"):
                sample_val = sample_val[:-2]
        self.samples[key].set(f'Ej: "{sample_val}"')

    def _update_all_samples(self):
        self._update_sample("card_code")
        for k in self.flags:
            self._toggle_field(k)

    def _apply_mapping(self):
        code_col = self.vars["card_code"].get()
        if not code_col or code_col not in self.df.columns:
            messagebox.showwarning("Atención", "Debe seleccionar una columna válida para el Código de Cliente.")
            return

        mapping = {k: self.vars[k].get() if (k == "card_code" or (self.flags[k].get() and self.vars[k].get() != "(Omitir / No actualizar)")) else "" for k in self.vars}
        flags = {k: bool(mapping[k]) for k in self.flags}

        self.on_apply(mapping, flags)
        self.destroy()


class BotClientesApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Bot de Clientes SAP B1 | RTN, Teléfonos, Correo, Direcciones y Zoho")
        self.root.geometry("1040x920")
        self.root.minsize(980, 800)

        self._setup_styles()

        self.engine = BotEngine()
        self.engine.on_pause_callback = self._on_engine_pause_change
        self.root.bind_all("<Escape>", lambda e: self.toggle_pause())
        self.root.bind_all("<F8>", lambda e: self.toggle_pause())
        self.worker_thread = None
        self.calibrating = False

        # Variables de estado
        self.excel_path_var = tk.StringVar(value="")
        self.selected_window_var = tk.StringVar(value="")
        self.window_status_var = tk.StringVar(value="Verificando ventanas...")
        self.speed_var = tk.StringVar(value="Normal (Recomendada)")
        self.status_bar_var = tk.StringVar(value="Listo. Cargue un archivo Excel para configurar el mapeo.")
        self.auto_recover_var = tk.BooleanVar(value=True)

        # Variables de etiquetas de mapeo visible
        self.lbl_mapped_code = tk.StringVar(value="Sin asignar")
        self.lbl_mapped_rtn = tk.StringVar(value="(Omitido)")
        self.lbl_mapped_tel = tk.StringVar(value="(Omitido)")
        self.lbl_mapped_mail = tk.StringVar(value="(Omitido)")
        self.lbl_mapped_dir = tk.StringVar(value="(Omitido)")

        # Variables de calibración de coordenadas
        self.coord_vars = {
            "btn_buscar": tk.StringVar(value="No calibrado"),
            "btn_crear": tk.StringVar(value="No calibrado"),
            "btn_confirmar_crear": tk.StringVar(value="No calibrado"),
            "card_code": tk.StringVar(value="No calibrado"),
            "btn_actualizar": tk.StringVar(value="No calibrado"),
            "barra_estado": tk.StringVar(value="No calibrado"),
            "menu_copiar_error": tk.StringVar(value="No calibrado"),
            "tab_general": tk.StringVar(value="No calibrado"),
            "tab_direcciones": tk.StringVar(value="No calibrado"),
            "rtn": tk.StringVar(value="No calibrado"),
            "telefono": tk.StringVar(value="No calibrado"),
            "movil": tk.StringVar(value="No calibrado"),
            "correo": tk.StringVar(value="No calibrado"),
            "activo": tk.StringVar(value="No calibrado"),
            "wbcustid": tk.StringVar(value="No calibrado"),
            "syncflag": tk.StringVar(value="No calibrado"),
            "definir_nuevo_factura": tk.StringVar(value="No calibrado"),
            "id_direccion": tk.StringVar(value="No calibrado"),
            "calle_numero": tk.StringVar(value="No calibrado"),
            "ciudad": tk.StringVar(value="No calibrado"),
            "btn_copiar_direccion": tk.StringVar(value="No calibrado"),
            "indicador_impuestos": tk.StringVar(value="No calibrado"),
        }

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
        cfg = self.engine.config
        key_map = {
            "btn_buscar": "btn_buscar_coord",
            "btn_crear": "btn_crear_coord",
            "btn_confirmar_crear": "btn_confirmar_crear_coord",
            "card_code": "card_code_coord",
            "btn_actualizar": "btn_actualizar_coord",
            "barra_estado": "barra_estado_coord",
            "menu_copiar_error": "menu_copiar_error_coord",
            "tab_general": "tab_general_coord",
            "tab_direcciones": "tab_direcciones_coord",
            "rtn": "rtn_coord",
            "telefono": "telefono_coord",
            "movil": "movil_coord",
            "correo": "correo_coord",
            "activo": "activo_coord",
            "wbcustid": "wbcustid_coord",
            "syncflag": "syncflag_coord",
            "definir_nuevo_factura": "definir_nuevo_factura_coord",
            "id_direccion": "id_direccion_coord",
            "calle_numero": "calle_numero_coord",
            "ciudad": "ciudad_coord",
            "btn_copiar_direccion": "btn_copiar_direccion_coord",
            "indicador_impuestos": "indicador_impuestos_coord",
        }
        for var_key, cfg_key in key_map.items():
            if cfg.get(cfg_key):
                x, y = cfg[cfg_key]
                self.coord_vars[var_key].set(f"X={x}, Y={y} ✓")

    def _create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ------------------- ENCABEZADO -------------------
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 6))

        lbl_title = ttk.Label(
            header_frame,
            text="Actualizador de Socios de Negocios (SAP B1 HANA)",
            style="Header.TLabel",
        )
        lbl_title.pack(anchor="w")

        lbl_sub = ttk.Label(
            header_frame,
            text="Automatización RPA para actualización masiva de RTN, Teléfonos, Correo, Direcciones y Zoho en 'Datos maestros socio de negocios'.",
            style="Subheader.TLabel",
        )
        lbl_sub.pack(anchor="w")

        ttk.Separator(main_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=(0, 6))

        # ------------------- TARJETA 1: ARCHIVO Y MAPEO DE COLUMNAS -------------------
        excel_card = ttk.LabelFrame(main_frame, text=" 1. Origen de Datos y Mapeo de Columnas ", padding="8")
        excel_card.pack(fill=tk.X, pady=(0, 6))

        file_row = ttk.Frame(excel_card)
        file_row.pack(fill=tk.X, pady=(0, 4))

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

        # Resumen visible
        map_summary_frame = ttk.Frame(excel_card)
        map_summary_frame.pack(fill=tk.X, pady=(2, 0))

        ttk.Label(map_summary_frame, text="Mapeo activo:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 6))

        ttk.Label(map_summary_frame, text="Código:").pack(side=tk.LEFT)
        ttk.Label(map_summary_frame, textvariable=self.lbl_mapped_code, style="MapBadge.TLabel").pack(side=tk.LEFT, padx=(2, 10))

        ttk.Label(map_summary_frame, text="RTN:").pack(side=tk.LEFT)
        ttk.Label(map_summary_frame, textvariable=self.lbl_mapped_rtn, style="MapBadge.TLabel").pack(side=tk.LEFT, padx=(2, 10))

        ttk.Label(map_summary_frame, text="Tel:").pack(side=tk.LEFT)
        ttk.Label(map_summary_frame, textvariable=self.lbl_mapped_tel, style="MapBadge.TLabel").pack(side=tk.LEFT, padx=(2, 10))

        ttk.Label(map_summary_frame, text="Correo:").pack(side=tk.LEFT)
        ttk.Label(map_summary_frame, textvariable=self.lbl_mapped_mail, style="MapBadge.TLabel").pack(side=tk.LEFT, padx=(2, 10))

        ttk.Label(map_summary_frame, text="Dirección:").pack(side=tk.LEFT)
        ttk.Label(map_summary_frame, textvariable=self.lbl_mapped_dir, style="MapBadge.TLabel").pack(side=tk.LEFT, padx=(2, 10))

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

        chk_recover = ttk.Checkbutton(
            status_row,
            text="🛡️ Auto-recuperar si falla al Actualizar (Crear ➔ Descartar ➔ Buscar y continuar)",
            variable=self.auto_recover_var,
            command=lambda: self.engine.config.update({"auto_recover_on_error": self.auto_recover_var.get()})
        )
        chk_recover.pack(side=tk.RIGHT)

        # ------------------- TARJETA 3: CALIBRACIÓN POR PESTAÑAS -------------------
        calib_card = ttk.LabelFrame(main_frame, text=" 3. Calibración de Pantalla en SAP ", padding="6")
        calib_card.pack(fill=tk.X, pady=(0, 6))

        nb_calib = ttk.Notebook(calib_card)
        nb_calib.pack(fill=tk.BOTH, expand=True, pady=(2, 4))

        # Pestaña A: Navegación y Control
        tab_nav = ttk.Frame(nb_calib, padding="6")
        nb_calib.add(tab_nav, text=" 🎮 Navegación y Control ")

        self._build_calib_row(tab_nav, "btn_buscar", "🔍 1. Botón 'Buscar' (Lupa barra SAP)", "btn_crear", "➕ 2. Botón 'Crear' (Añadir barra - recuperación)")
        self._build_calib_row(tab_nav, "btn_confirmar_crear", "✅ 3. Confirmar 'Crear nuevo' (Diálogo)", "card_code", "📝 4. Campo 'Código' (Cabecera SN)")
        self._build_calib_row(tab_nav, "btn_actualizar", "💾 5. Botón 'Buscar / Actualizar' (inferior)", "barra_estado", "📊 6. Barra de Estado (Clic derecho)")
        self._build_calib_row(tab_nav, "menu_copiar_error", "📋 7. Opción 'Copiar' (Menú clic derecho)", "tab_general", "📑 8. Pestaña 'General'")
        self._build_calib_row(tab_nav, "tab_direcciones", "📑 9. Pestaña 'Direcciones'", None, None)

        btn_wiz_nav = ttk.Button(
            tab_nav,
            text="🚀 Asistente: Calibrar Navegación y Control (en secuencia)",
            command=lambda: self.start_wizard_calibration("navegacion")
        )
        btn_wiz_nav.pack(anchor=tk.W, pady=(4, 0))

        # Pestaña B: Cabecera y General
        tab_gen = ttk.Frame(nb_calib, padding="6")
        nb_calib.add(tab_gen, text=" 📋 Cabecera, General y UDF ")

        self._build_calib_row(tab_gen, "rtn", "🆔 8. Campo 'RTN' (Cabecera)", "telefono", "📞 9. Campo 'Teléfono 1' (General)")
        self._build_calib_row(tab_gen, "movil", "📱 10. Campo 'Teléfono Móvil' (General)", "correo", "✉️ 11. Campo 'Correo Electrónico' (General)")
        self._build_calib_row(tab_gen, "activo", "🔘 12. Radio button 'Activo' (General)", "wbcustid", "🏷️ 13. Campo 'WBCUSTID' (UDF)")
        self._build_calib_row(tab_gen, "syncflag", "🚩 14. Campo 'SyncFlag' (UDF)", None, None)

        btn_wiz_gen = ttk.Button(
            tab_gen,
            text="🚀 Asistente: Calibrar Cabecera, General y UDF (en secuencia)",
            command=lambda: self.start_wizard_calibration("general")
        )
        btn_wiz_gen.pack(anchor=tk.W, pady=(4, 0))

        # Pestaña C: Direcciones (Destinatario ➔ Copiar >> ➔ Destino)
        tab_dir = ttk.Frame(nb_calib, padding="6")
        nb_calib.add(tab_dir, text=" 📍 Pestaña Direcciones ")

        self._build_calib_row(tab_dir, "definir_nuevo_factura", "📄 15. 'Definir nuevo' (Factura)", "id_direccion", "📍 16. Campo 'ID de Dirección'")
        self._build_calib_row(tab_dir, "calle_numero", "🏠 17. Campo 'Calle/ Número'", "ciudad", "🏙️ 18. Campo 'Ciudad'")
        self._build_calib_row(tab_dir, "btn_copiar_direccion", "⏩ 19. Botón 'Copiar >>' (a Destino)", "indicador_impuestos", "🏷️ 20. Campo 'Impuestos' (en Destino)")

        btn_wiz_dir = ttk.Button(
            tab_dir,
            text="🚀 Asistente: Calibrar Pestaña Direcciones (en secuencia)",
            command=lambda: self.start_wizard_calibration("direcciones")
        )
        btn_wiz_dir.pack(anchor=tk.W, pady=(4, 0))

        # Barra inferior de calibración
        calib_bottom = ttk.Frame(calib_card)
        calib_bottom.pack(fill=tk.X, pady=(2, 0))

        btn_wizard = ttk.Button(
            calib_bottom,
            text="🚀 Asistente Rápido: Calibrar en secuencia",
            command=self.start_wizard_calibration,
        )
        btn_wizard.pack(side=tk.LEFT, padx=(0, 10))

        btn_save = ttk.Button(calib_bottom, text="💾 Guardar Coordenadas", command=self._manual_save_coords)
        btn_save.pack(side=tk.LEFT, padx=(0, 10))

        btn_reset = ttk.Button(calib_bottom, text="🧹 Restablecer Pantalla SAP", command=self.reset_sap_screen)
        btn_reset.pack(side=tk.LEFT, padx=(0, 15))

        ttk.Label(calib_bottom, text="Velocidad RDP:").pack(side=tk.LEFT, padx=(0, 4))
        cbo_speed = ttk.Combobox(
            calib_bottom,
            textvariable=self.speed_var,
            values=["Rápida", "Normal (Recomendada)", "Lenta (Conexión inestable)"],
            state="readonly",
            width=22,
        )
        cbo_speed.pack(side=tk.LEFT)

        # ------------------- TARJETA 4: EJECUCIÓN Y RESULTADOS -------------------
        exec_card = ttk.LabelFrame(main_frame, text=" 4. Ejecución y Resultados ", padding="8")
        exec_card.pack(fill=tk.BOTH, expand=True, pady=(0, 4))

        btn_bar = ttk.Frame(exec_card)
        btn_bar.pack(fill=tk.X, pady=(0, 6))

        self.btn_test = ttk.Button(
            btn_bar,
            text="🧪 PROBAR CON 1 CLIENTE",
            command=self.start_single_test,
            style="ActionTest.TButton",
            width=26,
        )
        self.btn_test.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_start = ttk.Button(
            btn_bar,
            text="▶ INICIAR ACTUALIZACIÓN MASIVA",
            command=self.start_process,
            style="ActionStart.TButton",
            width=32,
        )
        self.btn_start.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_pause = ttk.Button(
            btn_bar,
            text="⏸ Pausar (F8 / ESC)",
            command=self.toggle_pause,
            state=tk.DISABLED,
            width=18,
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

        columns = ("num", "card_code", "rtn", "telefono", "correo", "direccion", "status", "message")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=6)

        self.tree.heading("num", text="#")
        self.tree.heading("card_code", text="Código")
        self.tree.heading("rtn", text="RTN")
        self.tree.heading("telefono", text="Teléfono")
        self.tree.heading("correo", text="Correo Electrónico")
        self.tree.heading("direccion", text="Dirección")
        self.tree.heading("status", text="Estado")
        self.tree.heading("message", text="Detalle")

        self.tree.column("num", width=35, anchor=tk.CENTER)
        self.tree.column("card_code", width=105, anchor=tk.CENTER)
        self.tree.column("rtn", width=115, anchor=tk.CENTER)
        self.tree.column("telefono", width=95, anchor=tk.CENTER)
        self.tree.column("correo", width=160, anchor=tk.W)
        self.tree.column("direccion", width=160, anchor=tk.W)
        self.tree.column("status", width=85, anchor=tk.CENTER)
        self.tree.column("message", width=250, anchor=tk.W)

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

    def _build_calib_row(self, parent, key1: str, label1: str, key2: Optional[str], label2: Optional[str]):
        row = ttk.Frame(parent)
        row.pack(fill=tk.X, pady=2)

        # Columna 1
        btn1 = ttk.Button(row, text=label1, width=35, command=lambda: self.start_calibration(key1))
        btn1.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(row, textvariable=self.coord_vars[key1], style="CoordBadge.TLabel", width=18).pack(side=tk.LEFT, padx=(0, 20))

        # Columna 2
        if key2 and label2:
            btn2 = ttk.Button(row, text=label2, width=35, command=lambda: self.start_calibration(key2))
            btn2.pack(side=tk.LEFT, padx=(0, 6))
            ttk.Label(row, textvariable=self.coord_vars[key2], style="CoordBadge.TLabel", width=18).pack(side=tk.LEFT)

    # ------------------- VENTANAS Y ARCHIVO -------------------

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
            title="Seleccionar archivo de clientes",
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
            self.lbl_row_count.config(text=f"✓ {total_rows} clientes", foreground="#15803d")
            self.status_bar_var.set(f"Archivo cargado: {os.path.basename(filepath)}. Abriendo asistente de mapeo...")

            self.open_mapping_dialog()

        except Exception as e:
            messagebox.showerror("Error al cargar archivo", f"No se pudo leer el archivo:\n{e}")

    def open_mapping_dialog(self):
        if self.engine.dataframe is None:
            messagebox.showwarning("Atención", "Cargue primero un archivo Excel.")
            return

        ClientColumnMappingDialog(
            parent=self.root,
            df=self.engine.dataframe,
            current_mapping=self.engine.column_mapping,
            current_flags=self.engine.update_flags,
            on_apply_callback=self._apply_mapped_columns,
        )

    def _apply_mapped_columns(self, mapping: dict, flags: dict):
        self.engine.column_mapping = mapping
        self.engine.update_flags = flags

        # Actualizar etiquetas de resumen
        self.lbl_mapped_code.set(mapping["card_code"] or "Sin asignar")
        self.lbl_mapped_rtn.set(mapping["rtn"] if flags.get("rtn") else "(Omitido)")
        self.lbl_mapped_tel.set(mapping["telefono"] if flags.get("telefono") else "(Omitido)")
        self.lbl_mapped_mail.set(mapping["correo"] if flags.get("correo") else "(Omitido)")
        self.lbl_mapped_dir.set(mapping["calle_numero"] if flags.get("calle_numero") else "(Omitido)")

        # Refrescar tabla de previsualización
        df = self.engine.dataframe
        self.tree.delete(*self.tree.get_children())

        for idx, row in df.iterrows():
            c_raw = row[mapping["card_code"]] if pd.notna(row[mapping["card_code"]]) else ""
            c_val = self.engine._clean_field_value("card_code", c_raw)

            r_raw = row[mapping["rtn"]] if flags.get("rtn") and mapping["rtn"] in df.columns and pd.notna(row[mapping["rtn"]]) else ""
            r_val = self.engine._clean_field_value("rtn", r_raw) if r_raw else "(Omitido)"

            t_raw = row[mapping["telefono"]] if flags.get("telefono") and mapping["telefono"] in df.columns and pd.notna(row[mapping["telefono"]]) else ""
            t_val = self.engine._clean_field_value("telefono", t_raw) if t_raw else "(Omitido)"

            m_raw = row[mapping["correo"]] if flags.get("correo") and mapping["correo"] in df.columns and pd.notna(row[mapping["correo"]]) else ""
            m_val = self.engine._clean_field_value("correo", m_raw) if m_raw else "(Omitido)"

            d_raw = row[mapping["calle_numero"]] if flags.get("calle_numero") and mapping["calle_numero"] in df.columns and pd.notna(row[mapping["calle_numero"]]) else ""
            d_val = self.engine._clean_field_value("calle_numero", d_raw) if d_raw else "(Omitido)"

            self.tree.insert("", tk.END, iid=f"row_{idx}", values=(idx + 1, c_val, r_val, t_val, m_val, d_val, "Listo", "En espera"))

        total_rows = len(df)
        self.prog_bar["maximum"] = total_rows
        self.prog_bar["value"] = 0
        self.lbl_progress_pct.config(text=f"0 / {total_rows} (0%)")
        self.status_bar_var.set("Mapeo aplicado. Verifique la calibración antes de iniciar.")

    # ------------------- CALIBRACIÓN -------------------

    def start_calibration(self, target_point: str, on_complete: Optional[callable] = None):
        if self.calibrating:
            return
        self.calibrating = True

        cfg_key = f"{target_point}_coord"
        label_name = target_point.replace("_", " ").title()

        calib_win = tk.Toplevel(self.root)
        calib_win.title(f"Calibrador: {label_name}")
        calib_win.geometry("480x170")
        calib_win.attributes("-topmost", True)
        calib_win.resizable(False, False)

        hints = {
            "btn_buscar": "🔍 Coloca el puntero sobre el botón 'Buscar'\n(icono de Lupa en la barra superior de SAP)...",
            "btn_crear": "➕ Coloca el puntero sobre el botón 'Crear / Añadir'\n(icono de hoja con signo + en barra de SAP)...",
            "btn_confirmar_crear": "✅ Coloca el puntero sobre el botón 'Sí' / 'OK'\ndel cuadro de diálogo que aparece al descartar cambios...",
            "card_code": "📝 Coloca el puntero sobre la casilla 'Código'\n(Código del Cliente en la cabecera de SAP)...",
            "btn_actualizar": "💾 Coloca el puntero sobre el botón 'Buscar / Actualizar'\n(esquina inferior izquierda de la ventana de SAP)...",
            "barra_estado": "📊 Coloca el puntero sobre la barra de estado inferior de SAP\n(franja inferior donde aparecen los mensajes en rojo o verde)...",
            "menu_copiar_error": "📋 Haz clic DERECHO manual en la barra de estado y coloca el puntero\nsobre la opción 'Copiar' / 'Copiar mensaje de error' del menú emergente...",
            "tab_general": "📑 Coloca el puntero sobre la pestaña 'General'\n(en la fila de pestañas al centro de la ventana)...",
            "rtn": "🆔 Coloca el puntero sobre la casilla 'RTN'\n(en la cabecera de datos maestros del socio de negocios)...",
            "telefono": "📞 Coloca el puntero sobre la casilla 'Teléfono 1'\n(dentro de la pestaña General)...",
            "movil": "📱 Coloca el puntero sobre la casilla 'Teléfono Móvil'\n(dentro de la pestaña General)...",
            "correo": "✉️ Coloca el puntero sobre la casilla 'Correo Electrónico'\n(dentro de la pestaña General)...",
            "activo": "🔘 Coloca el puntero sobre la opción / radio button 'Activo'\n(en la parte inferior de la pestaña General)...",
            "wbcustid": "🏷️ Coloca el puntero sobre la casilla 'WBCUSTID'\n(en el panel lateral derecho de Campos de Usuario UDF)...",
            "syncflag": "🚩 Coloca el puntero sobre la casilla 'SyncFlag'\n(en el panel lateral derecho de Campos de Usuario UDF)...",
            "tab_direcciones": "📍 Coloca el puntero sobre la pestaña 'Direcciones'\n(en la fila de pestañas al centro de la ventana)...",
            "definir_nuevo_factura": "📄 Coloca el puntero sobre 'Definir nuevo'\n(o la fila bajo 'Destinatario de factura' en el árbol izquierdo de direcciones)...",
            "id_direccion": "📍 Coloca el puntero sobre la casilla 'ID de dirección'\n(en el formulario derecho de la pestaña Direcciones)...",
            "calle_numero": "🏠 Coloca el puntero sobre la casilla 'Calle/ Número'\n(en el formulario derecho de la pestaña Direcciones)...",
            "ciudad": "🏙️ Coloca el puntero sobre la casilla 'Ciudad'\n(en el formulario derecho de la pestaña Direcciones)...",
            "btn_copiar_direccion": "⏩ Coloca el puntero sobre el botón 'Copiar >>'\n(botón central para clonar la dirección a Destino)...",
            "indicador_impuestos": "🏷️ Coloca el puntero sobre la casilla 'Indicador de impuestos'\n(en el formulario derecho tras copiar a Destino)...",
        }
        hint = hints.get(target_point, f"Coloca el puntero del mouse exactamente sobre:\n'{label_name}' en tu SAP...")

        lbl_info = ttk.Label(
            calib_win,
            text=hint,
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
                x, y = win32api.GetCursorPos()
                calib_win.destroy()
                self.calibrating = False

                self.engine.config[cfg_key] = (x, y)
                self.coord_vars[target_point].set(f"X={x}, Y={y} ✓")
                self.engine.save_coordinates_to_file()
                self.root.bell()

                if on_complete:
                    on_complete()
                else:
                    messagebox.showinfo("Calibración Exitosa", f"Se guardó la posición de:\n'{label_name}'\nX={x}, Y={y}")

        calib_win.after(1000, countdown_step, 4)

    def start_wizard_calibration(self, subset: Optional[str] = None):
        flags = self.engine.update_flags

        if subset == "navegacion":
            active_points = [
                "btn_buscar", "btn_crear", "btn_confirmar_crear", "card_code",
                "btn_actualizar", "barra_estado", "menu_copiar_error",
                "tab_general", "tab_direcciones"
            ]
            wizard_name = "Navegación y Control (9 puntos)"
        elif subset == "general":
            active_points = ["tab_general"]
            if flags.get("rtn", True): active_points.append("rtn")
            if flags.get("telefono", True): active_points.append("telefono")
            if flags.get("movil", True): active_points.append("movil")
            if flags.get("correo", True): active_points.append("correo")
            if flags.get("activo", True): active_points.append("activo")
            if flags.get("wbcustid", True): active_points.append("wbcustid")
            if flags.get("syncflag", True): active_points.append("syncflag")
            wizard_name = f"Cabecera, General y UDF ({len(active_points)} puntos)"
        elif subset == "direcciones":
            active_points = ["tab_direcciones", "definir_nuevo_factura"]
            if flags.get("id_direccion", True): active_points.append("id_direccion")
            if flags.get("calle_numero", True): active_points.append("calle_numero")
            if flags.get("ciudad", True): active_points.append("ciudad")
            active_points.append("btn_copiar_direccion")
            if flags.get("indicador_impuestos", True): active_points.append("indicador_impuestos")
            wizard_name = f"Pestaña Direcciones ({len(active_points)} puntos)"
        else:
            wizard_name = "Todos los Puntos Activos"
            active_points = [
                "btn_buscar", "btn_crear", "btn_confirmar_crear", "card_code",
                "btn_actualizar", "barra_estado", "menu_copiar_error"
            ]
            if any(flags.get(k) for k in ["rtn", "telefono", "movil", "correo", "activo"]):
                active_points.append("tab_general")
                if flags.get("rtn"): active_points.append("rtn")
                if flags.get("telefono"): active_points.append("telefono")
                if flags.get("movil"): active_points.append("movil")
                if flags.get("correo"): active_points.append("correo")
                if flags.get("activo"): active_points.append("activo")

            if flags.get("wbcustid"): active_points.append("wbcustid")
            if flags.get("syncflag"): active_points.append("syncflag")

            if any(flags.get(k) for k in ["id_direccion", "calle_numero", "ciudad", "indicador_impuestos"]):
                active_points.append("tab_direcciones")
                active_points.append("definir_nuevo_factura")
                if flags.get("id_direccion"): active_points.append("id_direccion")
                if flags.get("calle_numero"): active_points.append("calle_numero")
                if flags.get("ciudad"): active_points.append("ciudad")
                active_points.append("btn_copiar_direccion")
                if flags.get("indicador_impuestos"): active_points.append("indicador_impuestos")

        def run_step(step_idx):
            if step_idx < len(active_points):
                self.start_calibration(active_points[step_idx], on_complete=lambda: run_step(step_idx + 1))
            else:
                messagebox.showinfo("Asistente Completo", f"¡Todos los puntos de {wizard_name} fueron calibrados y guardados con éxito!")

        confirm = messagebox.askyesno(
            f"Asistente: {wizard_name}",
            f"El asistente te guiará para calibrar {len(active_points)} puntos en secuencia (5 seg por punto).\n\n"
            f"Coloca el cursor en tu SAP de la segunda pantalla según las instrucciones en cada paso.\n\n"
            f"¿Deseas iniciar ahora?"
        )
        if confirm:
            run_step(0)

    def reset_sap_screen(self):
        """Ejecuta el protocolo de limpieza y restablecimiento a Modo Buscar en SAP B1."""
        target = window_manager.find_target_window(self.selected_window_var.get().strip() or None)
        if not target:
            messagebox.showerror("SAP no detectado", "No se encontró la ventana de SAP Business One ni Escritorio Remoto abierto.")
            return
        hwnd, _ = target
        window_manager.bring_window_to_front(hwnd)
        time.sleep(0.2)
        try:
            self.engine._recover_and_reset_to_search()
            messagebox.showinfo("Pantalla Restablecida", "Se ejecutó la rutina de limpieza (Crear ➔ Descartar ➔ Buscar).\nSAP B1 ha vuelto a Modo Buscar limpiamente.")
        except Exception as e:
            messagebox.showerror("Error al restablecer", f"No se pudo completar el restablecimiento:\n{e}")

    def _manual_save_coords(self):
        saved = self.engine.save_coordinates_to_file()
        if saved:
            messagebox.showinfo("Guardado", "Coordenadas guardadas en disco en coordenadas_clientes.json.")
        else:
            messagebox.showerror("Error", "No se pudieron guardar las coordenadas en disco.")

    def _apply_settings(self):
        speed = self.speed_var.get()
        if "Rápida" in speed:
            self.engine.config["delay_after_find"] = 0.35
            self.engine.config["delay_after_search"] = 0.85
            self.engine.config["delay_after_tab_click"] = 0.30
            self.engine.config["delay_between_fields"] = 0.16
            self.engine.config["delay_after_save"] = 0.70
        elif "Lenta" in speed:
            self.engine.config["delay_after_find"] = 0.65
            self.engine.config["delay_after_search"] = 1.45
            self.engine.config["delay_after_tab_click"] = 0.55
            self.engine.config["delay_between_fields"] = 0.32
            self.engine.config["delay_after_save"] = 1.25
        else:
            self.engine.config["delay_after_find"] = 0.45
            self.engine.config["delay_after_search"] = 1.05
            self.engine.config["delay_after_tab_click"] = 0.40
            self.engine.config["delay_between_fields"] = 0.22
            self.engine.config["delay_after_save"] = 0.85

        self.engine.config["auto_recover_on_error"] = self.auto_recover_var.get()

    # ------------------- MODO DE PRUEBA: 1 CLIENTE -------------------

    def start_single_test(self):
        if not self.excel_path_var.get() or self.engine.dataframe is None:
            messagebox.showwarning("Atención", "Seleccione un archivo Excel antes de probar.")
            return

        cfg = self.engine.config
        if not cfg.get("card_code_coord"):
            messagebox.showwarning("Calibración Requerida", "Calibre al menos el campo 'Código de Cliente' (Cabecera).")
            return

        custom_title = self.selected_window_var.get().strip() or None
        target = window_manager.find_target_window(custom_title)
        if not target:
            messagebox.showerror("SAP no detectado", "No se encontró la ventana de SAP Business One ni Escritorio Remoto abierto.")
            return

        self.engine.config["custom_window_title"] = custom_title
        self._apply_settings()

        self.btn_test.config(state=tk.DISABLED)
        self.btn_start.config(state=tk.DISABLED)
        self.btn_pause.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)

        def worker():
            try:
                result = self.engine.run_single_test_client(
                    client_index=0,
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
        card_code = result.get("card_code", "")
        actualizados = ", ".join(result.get("actualizados", [])) or "Sin campos modificados"

        self.status_bar_var.set(f"Prueba completada para {card_code}. Revisa la pantalla de SAP.")

        resp = messagebox.askyesno(
            "Verificación de Prueba de 1 Cliente",
            f"✅ Se ejecutó la prueba con el cliente: {card_code}\n\n"
            f"Campos actualizados: {actualizados}\n\n"
            f"Por favor verifica en tu pantalla de SAP:\n"
            f"1. ¿Buscó el código de cliente en Modo Buscar?\n"
            f"2. ¿Digitó la información en los campos correspondientes?\n"
            f"3. ¿Se guardaron los cambios correctamente?\n\n"
            f"¿Todo fue correcto y deseas INICIAR la actualización masiva ahora?",
        )
        if resp:
            self.start_process()

    # ------------------- PROCESO MASIVO -------------------

    def start_process(self):
        if not self.excel_path_var.get() or self.engine.dataframe is None:
            messagebox.showwarning("Atención", "Seleccione un archivo Excel antes de iniciar.")
            return

        cfg = self.engine.config
        if not cfg.get("card_code_coord"):
            messagebox.showwarning("Calibración Requerida", "Es necesario calibrar al menos el campo 'Código de Cliente'.")
            return

        custom_title = self.selected_window_var.get().strip() or None
        target = window_manager.find_target_window(custom_title)
        if not target:
            messagebox.showerror("SAP no detectado", "No se encontró la ventana de SAP Business One ni Escritorio Remoto abierto.")
            return

        self.engine.config["custom_window_title"] = custom_title
        self._apply_settings()

        self.btn_test.config(state=tk.DISABLED)
        self.btn_start.config(state=tk.DISABLED)
        self.btn_pause.config(state=tk.NORMAL, text="⏸ Pausar (F8 / ESC)")
        self.btn_stop.config(state=tk.NORMAL)
        self.btn_export.config(state=tk.DISABLED)

        self.worker_thread = threading.Thread(target=self._run_worker, daemon=True)
        self.worker_thread.start()

    def _run_worker(self):
        try:
            self.engine.run_update_clients(
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
            self.root.after(0, lambda s=seconds_left: self.status_bar_var.set(f"⏳ Enfocando SAP... Iniciando en {s} segundos..."))
        else:
            self.root.after(0, lambda: self.status_bar_var.set("▶ Actualización de clientes en curso..."))

    def _on_row_progress(self, data: dict):
        def update_ui():
            idx = data["index"]
            total = data["total"]
            row_idx = data["row_idx"]
            status = data["status"]
            msg = data["message"]

            self.prog_bar["value"] = idx
            pct = int((idx / total) * 100) if total > 0 else 0
            self.lbl_progress_pct.config(text=f"{idx} / {total} ({pct}%)")

            item_iid = f"row_{row_idx}"
            if self.tree.exists(item_iid):
                cur_vals = list(self.tree.item(item_iid, "values"))
                cur_vals[6] = status
                cur_vals[7] = msg
                self.tree.item(item_iid, values=cur_vals, tags=(status,))
                self.tree.see(item_iid)

            self.status_bar_var.set(f"[{idx}/{total}] {data['card_code']} - {msg}")

        self.root.after(0, update_ui)

    def _on_sync_finished(self, stats: dict):
        def finish_ui():
            self._reset_controls()
            self.btn_export.config(state=tk.NORMAL)

            if stats.get("aborted"):
                self.status_bar_var.set(f"⏹ Proceso cancelado o detenido por el usuario. Exitosos: {stats['success']}, Fallidos: {stats['failed']}")
                messagebox.showwarning("Proceso Detenido", f"El proceso se detuvo.\nClientes actualizados: {stats['success']}\nCon error: {stats['failed']}")
            else:
                self.status_bar_var.set(f"✅ Proceso finalizado. Total: {stats['total']}, Exitosos: {stats['success']}, Fallidos: {stats['failed']}")
                messagebox.showinfo("Proceso Completado", f"Actualización finalizada con éxito.\n\nTotal procesados: {stats['processed']}\nExitosos: {stats['success']}\nFallidos / Omitidos: {stats['failed']}")

        self.root.after(0, finish_ui)

    def _on_engine_pause_change(self, is_paused: bool):
        def update_pause():
            if is_paused:
                self.btn_pause.config(text="▶ Reanudar (F8 / ESC)")
                self.status_bar_var.set("⏸ Bot en PAUSA. Presiona F8 o haz clic en Reanudar.")
            else:
                self.btn_pause.config(text="⏸ Pausar (F8 / ESC)")
                self.status_bar_var.set("▶ Reanudando actualización de clientes...")
        self.root.after(0, update_pause)

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
        if self.engine.is_running:
            self.engine.stop()
            self.status_bar_var.set("Deteniendo proceso de actualización...")

    def _reset_controls(self):
        self.btn_test.config(state=tk.NORMAL)
        self.btn_start.config(state=tk.NORMAL)
        self.btn_pause.config(state=tk.DISABLED, text="⏸ Pausar (F8 / ESC)")
        self.btn_stop.config(state=tk.DISABLED)

    def export_results(self):
        filepath = filedialog.asksaveasfilename(
            title="Exportar resultados de clientes",
            defaultextension=".xlsx",
            filetypes=[("Excel (.xlsx)", "*.xlsx"), ("CSV (.csv)", "*.csv")],
        )
        if not filepath:
            return
        try:
            records = []
            for item_id in self.tree.get_children():
                vals = self.tree.item(item_id, "values")
                records.append({
                    "#": vals[0],
                    "Codigo_Cliente": vals[1],
                    "RTN": vals[2],
                    "Telefono": vals[3],
                    "Correo": vals[4],
                    "Direccion": vals[5],
                    "Estado": vals[6],
                    "Detalle": vals[7],
                })
            df_out = pd.DataFrame(records)
            if filepath.lower().endswith(".csv"):
                df_out.to_csv(filepath, index=False, encoding="utf-8-sig")
            else:
                df_out.to_excel(filepath, index=False)
            messagebox.showinfo("Exportación Exitosa", f"Resultados exportados a:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Error al Exportar", f"No se pudo guardar el archivo:\n{e}")


def main():
    root = tk.Tk()
    app = BotClientesApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
