"""
app_gui.py
Interfaz Gráfica de Usuario (Tkinter) para el Bot de SAP Business One (HANA).
Soporta dos modos:
  1. Modo Sincronización: Digitación de WBCUSTID y SyncFlag = 'T'.
  2. Modo Eliminación: Búsqueda y eliminación de socios de negocio duplicados.
Lazarus & Lazarus.
"""

import os
import sys
import time
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import pyautogui

import window_manager
from bot_engine import BotEngine, SAPWindowNotFoundError


class BotClientesApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Bot de Gestión de Clientes | SAP B1 HANA ➔ Zoho CRM")
        self.root.geometry("940x820")
        self.root.minsize(880, 720)

        self._setup_styles()

        self.engine = BotEngine()
        self.worker_thread = None
        self.calibrating = False

        # Variables de estado
        self.current_mode_var = tk.StringVar(value="sync")  # 'sync' o 'delete'
        self.excel_path_var = tk.StringVar(value="")
        self.selected_window_var = tk.StringVar(value="")
        self.window_status_var = tk.StringVar(value="Verificando ventanas...")
        self.sap_col_var = tk.StringVar(value="")
        self.zoho_col_var = tk.StringVar(value="")
        self.wbcustid_coord_var = tk.StringVar(value="No calibrado (Haga clic en 'Calibrar' o use Tabs)")
        self.syncflag_mode_var = tk.StringVar(value="Avanzar 1 Tab desde WBCUSTID (Recomendado)")
        self.syncflag_coord_var = tk.StringVar(value="No calibrado")
        self.syncflag_val_var = tk.StringVar(value="T")
        self.speed_var = tk.StringVar(value="Normal (Recomendada)")
        self.save_action_var = tk.StringVar(value="Alt + A (Atajo SAP)")
        self.status_bar_var = tk.StringVar(value="Listo para iniciar. Seleccione el modo y cargue un archivo.")

        # Variables Modo Eliminación
        self.delete_method_var = tk.StringVar(value="Menú Datos SAP (Alt + D ➔ Eliminar) [Recomendado]")
        self.delete_coord_var = tk.StringVar(value="No calibrado")

        self._create_widgets()
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
        style.configure("WarnBanner.TLabel", font=("Segoe UI", 9, "bold"), foreground="#9a3412", background="#ffedd5")
        style.configure("ActionStart.TButton", font=("Segoe UI", 10, "bold"))
        style.configure("ActionDelete.TButton", font=("Segoe UI", 10, "bold"))
        style.configure("ActionStop.TButton", font=("Segoe UI", 10, "bold"))

    def _create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ------------------- ENCABEZADO -------------------
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 6))

        lbl_title = ttk.Label(
            header_frame,
            text="Bot de Clientes SAP Business One (HANA)",
            style="Header.TLabel",
        )
        lbl_title.pack(anchor="w")

        lbl_sub = ttk.Label(
            header_frame,
            text="Automatización RPA: Sincronización con Zoho CRM y Eliminación de Duplicados en SAP.",
            style="Subheader.TLabel",
        )
        lbl_sub.pack(anchor="w")

        # ------------------- SELECTOR DE MODO -------------------
        mode_frame = ttk.LabelFrame(main_frame, text=" Modo de Operación ", padding="6")
        mode_frame.pack(fill=tk.X, pady=(4, 8))

        rb_sync = ttk.Radiobutton(
            mode_frame,
            text="🔄 Modo 1: Sincronizar Zoho CRM (Digitar WBCUSTID y SyncFlag = 'T')",
            value="sync",
            variable=self.current_mode_var,
            command=self._on_mode_changed,
        )
        rb_sync.pack(side=tk.LEFT, padx=(10, 30))

        rb_delete = ttk.Radiobutton(
            mode_frame,
            text="🗑️ Modo 2: Eliminar Clientes Duplicados en SAP (Búsqueda ➔ Clic derecho / Menú ➔ Eliminar)",
            value="delete",
            variable=self.current_mode_var,
            command=self._on_mode_changed,
        )
        rb_delete.pack(side=tk.LEFT)

        # Banner de advertencia dinámico
        self.banner_frame = ttk.Frame(main_frame)
        self.lbl_banner = ttk.Label(
            self.banner_frame,
            text="",
            style="WarnBanner.TLabel",
            padding="6",
            wraplength=880,
            justify=tk.LEFT,
        )
        self.lbl_banner.pack(fill=tk.X)

        # ------------------- TARJETA 1: ARCHIVO EXCEL -------------------
        excel_card = ttk.LabelFrame(main_frame, text=" 1. Origen de Datos (Excel / CSV) ", padding="8")
        excel_card.pack(fill=tk.X, pady=(0, 8))

        file_row = ttk.Frame(excel_card)
        file_row.pack(fill=tk.X, pady=(0, 4))

        ttk.Label(file_row, text="Archivo:", width=8).pack(side=tk.LEFT)
        ent_file = ttk.Entry(file_row, textvariable=self.excel_path_var, state="readonly")
        ent_file.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        btn_browse = ttk.Button(file_row, text="📁 Examinar...", command=self.browse_file)
        btn_browse.pack(side=tk.RIGHT)

        self.col_row = ttk.Frame(excel_card)
        self.col_row.pack(fill=tk.X, pady=(4, 0))

        ttk.Label(self.col_row, text="Col. Código SAP:").pack(side=tk.LEFT, padx=(0, 4))
        self.cbo_sap_col = ttk.Combobox(self.col_row, textvariable=self.sap_col_var, state="readonly", width=18)
        self.cbo_sap_col.pack(side=tk.LEFT, padx=(0, 16))

        self.lbl_zoho_col = ttk.Label(self.col_row, text="Col. WBCUSTID / Zoho:")
        self.lbl_zoho_col.pack(side=tk.LEFT, padx=(0, 4))
        self.cbo_zoho_col = ttk.Combobox(self.col_row, textvariable=self.zoho_col_var, state="readonly", width=18)
        self.cbo_zoho_col.pack(side=tk.LEFT, padx=(0, 16))

        self.lbl_row_count = ttk.Label(self.col_row, text="0 registros cargados", foreground="#64748b")
        self.lbl_row_count.pack(side=tk.LEFT)

        # ------------------- TARJETA 2: DETECCIÓN DE VENTANA SAP / RDP -------------------
        win_card = ttk.LabelFrame(main_frame, text=" 2. Ventana de SAP / Escritorio Remoto ", padding="8")
        win_card.pack(fill=tk.X, pady=(0, 8))

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
        status_row.pack(fill=tk.X, pady=(4, 0))
        self.lbl_win_status = ttk.Label(status_row, textvariable=self.window_status_var, style="StatusOk.TLabel")
        self.lbl_win_status.pack(side=tk.LEFT)

        # ------------------- TARJETA 3: PARÁMETROS ESPECÍFICOS SEGÚN MODO -------------------
        self.config_card = ttk.LabelFrame(main_frame, text=" 3. Parámetros de Operación ", padding="8")
        self.config_card.pack(fill=tk.X, pady=(0, 8))

        # Contenedor para controles de Modo Sincronización
        self.frame_sync_params = ttk.Frame(self.config_card)
        self.frame_sync_params.pack(fill=tk.X)

        row_wbcust = ttk.Frame(self.frame_sync_params)
        row_wbcust.pack(fill=tk.X, pady=(0, 4))
        btn_calib_wb = ttk.Button(
            row_wbcust,
            text="🎯 Calibrar Clic en WBCUSTID",
            command=lambda: self.start_calibration("wbcustid"),
        )
        btn_calib_wb.pack(side=tk.LEFT, padx=(0, 10))
        lbl_coord_wb = ttk.Label(row_wbcust, textvariable=self.wbcustid_coord_var, foreground="#0369a1")
        lbl_coord_wb.pack(side=tk.LEFT, fill=tk.X, expand=True)
        btn_clear_wb = ttk.Button(row_wbcust, text="Limpiar Clic", command=self.clear_wbcustid_calibration)
        btn_clear_wb.pack(side=tk.RIGHT)

        row_syncflag = ttk.Frame(self.frame_sync_params)
        row_syncflag.pack(fill=tk.X, pady=(4, 4))
        ttk.Label(row_syncflag, text="Llegada a SyncFlag:").pack(side=tk.LEFT, padx=(0, 5))
        cbo_sync_mode = ttk.Combobox(
            row_syncflag,
            textvariable=self.syncflag_mode_var,
            values=[
                "Avanzar 1 Tab desde WBCUSTID (Recomendado)",
                "Avanzar 2 Tabs desde WBCUSTID",
                "Calibrar Clic independiente en SyncFlag",
            ],
            state="readonly",
            width=38,
        )
        cbo_sync_mode.pack(side=tk.LEFT, padx=(0, 10))
        btn_calib_sf = ttk.Button(
            row_syncflag,
            text="🎯 Calibrar Clic SyncFlag",
            command=lambda: self.start_calibration("syncflag"),
        )
        btn_calib_sf.pack(side=tk.LEFT, padx=(0, 8))
        lbl_sf_coord = ttk.Label(row_syncflag, textvariable=self.syncflag_coord_var, foreground="#64748b")
        lbl_sf_coord.pack(side=tk.LEFT)

        row_sync_sub = ttk.Frame(self.frame_sync_params)
        row_sync_sub.pack(fill=tk.X, pady=(4, 0))
        ttk.Label(row_sync_sub, text="Valor SyncFlag:").pack(side=tk.LEFT, padx=(0, 4))
        ent_sf_val = ttk.Entry(row_sync_sub, textvariable=self.syncflag_val_var, width=5, justify=tk.CENTER)
        ent_sf_val.pack(side=tk.LEFT, padx=(0, 15))
        ttk.Label(row_sync_sub, text="Acción al Guardar:").pack(side=tk.LEFT, padx=(0, 4))
        cbo_save = ttk.Combobox(
            row_sync_sub,
            textvariable=self.save_action_var,
            values=["Alt + A (Atajo SAP)", "Tecla Enter"],
            state="readonly",
            width=18,
        )
        cbo_save.pack(side=tk.LEFT, padx=(0, 15))

        # Contenedor para controles de Modo Eliminación
        self.frame_delete_params = ttk.Frame(self.config_card)

        row_del1 = ttk.Frame(self.frame_delete_params)
        row_del1.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(row_del1, text="Método de Eliminación:").pack(side=tk.LEFT, padx=(0, 5))
        cbo_del_method = ttk.Combobox(
            row_del1,
            textvariable=self.delete_method_var,
            values=[
                "Menú Datos SAP (Alt + D ➔ Eliminar) [Recomendado]",
                "Clic derecho en cabecera de socio de negocios",
            ],
            state="readonly",
            width=46,
        )
        cbo_del_method.pack(side=tk.LEFT, padx=(0, 10))

        btn_calib_del = ttk.Button(
            row_del1,
            text="🎯 Calibrar Clic Derecho en Cabecera",
            command=lambda: self.start_calibration("delete_header"),
        )
        btn_calib_del.pack(side=tk.LEFT, padx=(0, 8))
        lbl_del_coord = ttk.Label(row_del1, textvariable=self.delete_coord_var, foreground="#64748b")
        lbl_del_coord.pack(side=tk.LEFT)

        # Configuración común: Velocidad RDP
        row_speed = ttk.Frame(self.config_card)
        row_speed.pack(fill=tk.X, pady=(6, 0))
        ttk.Label(row_speed, text="Velocidad Escritorio Remoto:").pack(side=tk.LEFT, padx=(0, 4))
        cbo_speed = ttk.Combobox(
            row_speed,
            textvariable=self.speed_var,
            values=["Rápida", "Normal (Recomendada)", "Lenta (Conexión inestable)"],
            state="readonly",
            width=22,
        )
        cbo_speed.pack(side=tk.LEFT)

        # ------------------- TARJETA 4: PROGRESO Y CONTROLES -------------------
        exec_card = ttk.LabelFrame(main_frame, text=" 4. Ejecución y Resultados ", padding="8")
        exec_card.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        btn_bar = ttk.Frame(exec_card)
        btn_bar.pack(fill=tk.X, pady=(0, 8))

        self.btn_start = ttk.Button(
            btn_bar,
            text="▶ INICIAR DIGITACIÓN",
            command=self.start_process,
            style="ActionStart.TButton",
            width=28,
        )
        self.btn_start.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_pause = ttk.Button(
            btn_bar,
            text="⏸ Pausar",
            command=self.toggle_pause,
            state=tk.DISABLED,
            width=12,
        )
        self.btn_pause.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_stop = ttk.Button(
            btn_bar,
            text="⏹ DETENER (ESC)",
            command=self.stop_process,
            state=tk.DISABLED,
            style="ActionStop.TButton",
            width=16,
        )
        self.btn_stop.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_export = ttk.Button(
            btn_bar,
            text="💾 Exportar Resultados",
            command=self.export_results,
            state=tk.DISABLED,
        )
        self.btn_export.pack(side=tk.RIGHT)

        # Barra de progreso
        prog_frame = ttk.Frame(exec_card)
        prog_frame.pack(fill=tk.X, pady=(0, 6))

        self.prog_bar = ttk.Progressbar(prog_frame, orient=tk.HORIZONTAL, mode="determinate")
        self.prog_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        self.lbl_progress_pct = ttk.Label(prog_frame, text="0 / 0 (0%)", width=14)
        self.lbl_progress_pct.pack(side=tk.RIGHT)

        # Tabla de resultados dinámica
        tree_frame = ttk.Frame(exec_card)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(tree_frame, show="headings", height=8)
        self._setup_treeview_columns()

        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.tag_configure("OK", foreground="#15803d")
        self.tree.tag_configure("ELIMINADO", foreground="#c2410c")
        self.tree.tag_configure("ERROR", foreground="#b91c1c")
        self.tree.tag_configure("PROCESANDO", foreground="#2563eb")
        self.tree.tag_configure("BUSCANDO", foreground="#2563eb")

        # Barra de estado inferior
        status_bar = ttk.Frame(main_frame)
        status_bar.pack(fill=tk.X, pady=(3, 0))
        lbl_bot_status = ttk.Label(status_bar, textvariable=self.status_bar_var, relief=tk.SUNKEN, anchor="w", padding="3")
        lbl_bot_status.pack(fill=tk.X)

        self._on_mode_changed()

    def _setup_treeview_columns(self):
        mode = self.current_mode_var.get()
        self.tree.delete(*self.tree.get_children())

        if mode == "sync":
            self.tree["columns"] = ("num", "sap_code", "wbcustid", "syncflag", "status", "message")
            self.tree.heading("num", text="#")
            self.tree.heading("sap_code", text="Código SAP")
            self.tree.heading("wbcustid", text="WBCUSTID (Zoho)")
            self.tree.heading("syncflag", text="SyncFlag")
            self.tree.heading("status", text="Estado")
            self.tree.heading("message", text="Detalle de Operación")

            self.tree.column("num", width=40, anchor=tk.CENTER)
            self.tree.column("sap_code", width=110, anchor=tk.CENTER)
            self.tree.column("wbcustid", width=160, anchor=tk.CENTER)
            self.tree.column("syncflag", width=70, anchor=tk.CENTER)
            self.tree.column("status", width=90, anchor=tk.CENTER)
            self.tree.column("message", width=340, anchor=tk.W)
        else:  # Modo Eliminación
            self.tree["columns"] = ("num", "sap_code", "status", "message")
            self.tree.heading("num", text="#")
            self.tree.heading("sap_code", text="Código SAP a Eliminar")
            self.tree.heading("status", text="Estado")
            self.tree.heading("message", text="Resultado en SAP")

            self.tree.column("num", width=45, anchor=tk.CENTER)
            self.tree.column("sap_code", width=160, anchor=tk.CENTER)
            self.tree.column("status", width=110, anchor=tk.CENTER)
            self.tree.column("message", width=450, anchor=tk.W)

    def _on_mode_changed(self):
        mode = self.current_mode_var.get()
        self._setup_treeview_columns()

        if mode == "sync":
            self.banner_frame.pack_forget()
            self.frame_delete_params.pack_forget()
            self.frame_sync_params.pack(fill=tk.X)
            self.lbl_zoho_col.pack(side=tk.LEFT, padx=(0, 4))
            self.cbo_zoho_col.pack(side=tk.LEFT, padx=(0, 16))
            self.btn_start.config(text="▶ INICIAR DIGITACIÓN ZOHO")
            self.status_bar_var.set("Modo Sincronización Zoho activo. Cargue archivo con códigos y WBCUSTID.")
        else:
            self.banner_frame.pack(fill=tk.X, pady=(0, 6), before=self.col_row.master)
            self.lbl_banner.config(
                text="⚠️ ADVERTENCIA: En este modo el bot buscará cada cliente y ejecutará 'Eliminar' en SAP B1.\n"
                     "SAP solo permite eliminar aquellos clientes que no posean historial de transacciones (facturas, pedidos, etc.)."
            )
            self.frame_sync_params.pack_forget()
            self.frame_delete_params.pack(fill=tk.X)
            self.lbl_zoho_col.pack_forget()
            self.cbo_zoho_col.pack_forget()
            self.btn_start.config(text="🗑️ INICIAR ELIMINACIÓN EN SAP")
            self.status_bar_var.set("Modo Eliminación activo. Cargue archivo con la lista de códigos a eliminar.")

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
            df, sap_col, zoho_col = self.engine.load_file(filepath)
            self.excel_path_var.set(filepath)

            cols = list(df.columns)
            self.cbo_sap_col["values"] = cols
            self.cbo_zoho_col["values"] = cols
            self.sap_col_var.set(sap_col)
            self.zoho_col_var.set(zoho_col)

            total_rows = len(df)
            self.lbl_row_count.config(text=f"✓ {total_rows} registros detectados", foreground="#15803d")
            self.status_bar_var.set(f"Archivo cargado: {os.path.basename(filepath)} ({total_rows} filas)")

            mode = self.current_mode_var.get()
            self.tree.delete(*self.tree.get_children())
            for idx, row in df.iterrows():
                s_val = str(row[sap_col]) if pd.notna(row[sap_col]) else ""
                if mode == "sync":
                    z_val = str(row[zoho_col]) if pd.notna(row[zoho_col]) else ""
                    if z_val.endswith(".0"):
                        z_val = z_val[:-2]
                    self.tree.insert("", tk.END, iid=f"row_{idx}", values=(idx + 1, s_val, z_val, "T", "Pendiente", "En espera"))
                else:
                    self.tree.insert("", tk.END, iid=f"row_{idx}", values=(idx + 1, s_val, "Pendiente", "En espera para eliminar"))

            self.prog_bar["maximum"] = total_rows
            self.prog_bar["value"] = 0
            self.lbl_progress_pct.config(text=f"0 / {total_rows} (0%)")

        except Exception as e:
            messagebox.showerror("Error al cargar archivo", f"No se pudo leer el archivo:\n{e}")

    def start_calibration(self, target_field: str = "wbcustid"):
        if self.calibrating:
            return
        self.calibrating = True

        field_names = {
            "wbcustid": "WBCUSTID",
            "syncflag": "SyncFlag",
            "delete_header": "Cabecera de Socio de Negocios (para Clic Derecho)",
        }
        name = field_names.get(target_field, "Campo")

        calib_win = tk.Toplevel(self.root)
        calib_win.title(f"Calibrador: {name}")
        calib_win.geometry("420x150")
        calib_win.attributes("-topmost", True)
        calib_win.resizable(False, False)

        lbl_info = ttk.Label(
            calib_win,
            text=f"Coloca el puntero del mouse sobre:\n'{name}' en tu pantalla de SAP...",
            font=("Segoe UI", 10),
            justify=tk.CENTER,
            padding=10,
        )
        lbl_info.pack()

        lbl_countdown = ttk.Label(calib_win, text="5", font=("Segoe UI", 26, "bold"), foreground="#e11d48")
        lbl_countdown.pack()

        def countdown_step(remaining):
            if remaining > 0:
                lbl_countdown.config(text=str(remaining))
                calib_win.after(1000, countdown_step, remaining - 1)
            else:
                x, y = pyautogui.position()
                calib_win.destroy()
                self.calibrating = False

                if target_field == "wbcustid":
                    self.engine.config["zoho_field_coord"] = (x, y)
                    self.wbcustid_coord_var.set(f"Posición: X={x}, Y={y}")
                elif target_field == "syncflag":
                    self.engine.config["syncflag_coord"] = (x, y)
                    self.syncflag_coord_var.set(f"Posición: X={x}, Y={y}")
                    self.syncflag_mode_var.set("Calibrar Clic independiente en SyncFlag")
                else:  # delete_header
                    self.engine.config["delete_click_coord"] = (x, y)
                    self.delete_coord_var.set(f"Posición Cabecera: X={x}, Y={y}")
                    self.delete_method_var.set("Clic derecho en cabecera de socio de negocios")

                messagebox.showinfo(
                    "Calibración Exitosa",
                    f"Se ha guardado la posición de:\n'{name}'\nX = {x}, Y = {y}",
                )

        calib_win.after(1000, countdown_step, 4)

    def clear_wbcustid_calibration(self):
        self.engine.config["zoho_field_coord"] = None
        self.wbcustid_coord_var.set("No calibrado (Haga clic en 'Calibrar' o use Tabs)")

    def _apply_settings(self):
        speed = self.speed_var.get()
        if "Rápida" in speed:
            self.engine.config["delay_after_find"] = 0.3
            self.engine.config["delay_after_search"] = 0.6
            self.engine.config["delay_after_field_focus"] = 0.15
            self.engine.config["delay_between_fields"] = 0.15
            self.engine.config["delay_after_save"] = 0.5
            self.engine.config["delay_after_delete_action"] = 0.4
            self.engine.config["delay_after_confirm_delete"] = 0.6
        elif "Lenta" in speed:
            self.engine.config["delay_after_find"] = 0.6
            self.engine.config["delay_after_search"] = 1.3
            self.engine.config["delay_after_field_focus"] = 0.3
            self.engine.config["delay_between_fields"] = 0.3
            self.engine.config["delay_after_save"] = 1.0
            self.engine.config["delay_after_delete_action"] = 0.7
            self.engine.config["delay_after_confirm_delete"] = 1.2
        else:  # Normal
            self.engine.config["delay_after_find"] = 0.4
            self.engine.config["delay_after_search"] = 0.9
            self.engine.config["delay_after_field_focus"] = 0.2
            self.engine.config["delay_between_fields"] = 0.2
            self.engine.config["delay_after_save"] = 0.7
            self.engine.config["delay_after_delete_action"] = 0.5
            self.engine.config["delay_after_confirm_delete"] = 0.8

        action = self.save_action_var.get()
        self.engine.config["action_save"] = "enter" if "Enter" in action else "alt_a"

        sync_mode = self.syncflag_mode_var.get()
        if "Clic" in sync_mode:
            self.engine.config["syncflag_mode"] = "click"
        elif "2 Tabs" in sync_mode:
            self.engine.config["syncflag_mode"] = "tab"
            self.engine.config["syncflag_tab_count"] = 2
        else:
            self.engine.config["syncflag_mode"] = "tab"
            self.engine.config["syncflag_tab_count"] = 1

        self.engine.config["syncflag_value"] = self.syncflag_val_var.get().strip() or "T"

        # Configuración de eliminación
        del_method = self.delete_method_var.get()
        if "Clic derecho" in del_method:
            self.engine.config["delete_method"] = "right_click"
        else:
            self.engine.config["delete_method"] = "menu_alt_d"

    def start_process(self):
        if not self.excel_path_var.get() or self.engine.dataframe is None:
            messagebox.showwarning("Atención", "Por favor seleccione un archivo Excel antes de iniciar.")
            return

        mode = self.current_mode_var.get()
        if mode == "delete":
            confirm = messagebox.askyesno(
                "Confirmación de Eliminación Masiva",
                "⚠️ ¿ESTÁ SEGURO DE INICIAR LA ELIMINACIÓN?\n\n"
                "El bot buscará cada cliente y ejecutará la orden de eliminar en SAP B1.\n\n"
                "¿Desea proceder?",
                icon="warning",
            )
            if not confirm:
                return

        # Validación obligatoria de ventana SAP
        custom_title = self.selected_window_var.get().strip() or None
        target = window_manager.find_target_window(custom_title)
        if not target:
            messagebox.showerror(
                "Error: SAP / Escritorio Remoto no detectado",
                "⚠️ NO SE ENCONTRÓ LA VENTANA DE SAP BUSINESS ONE NI DE ESCRITORIO REMOTO.\n\n"
                "Para poder ejecutar:\n"
                "1. Abre la conexión de Escritorio Remoto (182.160.29.90).\n"
                "2. Inicia sesión en SAP Business One con tu usuario.\n"
                "3. Abre la pantalla 'Datos maestros de socio de negocios'.\n"
                "4. Vuelve a hacer clic en Iniciar.",
            )
            return

        self.engine.config["custom_window_title"] = custom_title
        self.engine.sap_col = self.sap_col_var.get()
        self.engine.zoho_col = self.zoho_col_var.get()
        self._apply_settings()

        self.btn_start.config(state=tk.DISABLED)
        self.btn_pause.config(state=tk.NORMAL, text="⏸ Pausar")
        self.btn_stop.config(state=tk.NORMAL)
        self.btn_export.config(state=tk.DISABLED)

        self.worker_thread = threading.Thread(target=self._run_worker, daemon=True)
        self.worker_thread.start()

    def _run_worker(self):
        mode = self.current_mode_var.get()
        try:
            if mode == "sync":
                self.engine.run_sync(
                    progress_callback=self._on_row_progress_sync,
                    on_countdown_tick=self._on_countdown,
                    completion_callback=self._on_sync_finished,
                )
            else:
                self.engine.run_delete_clients(
                    progress_callback=self._on_row_progress_delete,
                    on_countdown_tick=self._on_countdown,
                    completion_callback=self._on_delete_finished,
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
                    f"⏳ Enfocando SAP... Iniciando ejecución en {s} segundos..."
                ),
            )
        else:
            self.root.after(0, lambda: self.status_bar_var.set("▶ Proceso en curso en SAP..."))

    def _on_row_progress_sync(self, data: dict):
        def update_ui():
            idx = data["index"]
            total = data["total"]
            row_idx = data["row_idx"]
            status = data["status"]
            msg = data["message"]
            sf_val = data.get("syncflag", "T")

            pct = int((idx / total) * 100) if total > 0 else 0
            self.prog_bar["value"] = idx
            self.lbl_progress_pct.config(text=f"{idx} / {total} ({pct}%)")
            self.status_bar_var.set(f"[{idx}/{total}] Cliente {data['card_code']} ➔ WBCUSTID={data['zoho_id']} | SyncFlag={sf_val}")

            item_id = f"row_{row_idx}"
            if self.tree.exists(item_id):
                self.tree.item(
                    item_id,
                    values=(idx, data["card_code"], data["zoho_id"], sf_val, status, msg),
                    tags=(status,),
                )
                self.tree.see(item_id)

        self.root.after(0, update_ui)

    def _on_row_progress_delete(self, data: dict):
        def update_ui():
            idx = data["index"]
            total = data["total"]
            row_idx = data["row_idx"]
            status = data["status"]
            msg = data["message"]

            pct = int((idx / total) * 100) if total > 0 else 0
            self.prog_bar["value"] = idx
            self.lbl_progress_pct.config(text=f"{idx} / {total} ({pct}%)")
            self.status_bar_var.set(f"[{idx}/{total}] Eliminando cliente {data['card_code']}...")

            item_id = f"row_{row_idx}"
            if self.tree.exists(item_id):
                self.tree.item(
                    item_id,
                    values=(idx, data["card_code"], status, msg),
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
                messagebox.showinfo("Proceso Detenido", f"El proceso fue detenido.\n\nClientes actualizados: {stats['success']}/{stats['total']}")
            else:
                self.status_bar_var.set(f"✓ Finalizado: {stats['success']} actualizados, {stats['failed']} errores.")
                messagebox.showinfo(
                    "Sincronización Completada",
                    f"¡Proceso finalizado exitosamente!\n\n"
                    f"Total registros: {stats['total']}\n"
                    f"Actualizados con éxito: {stats['success']}\n"
                    f"Errores: {stats['failed']}\n\n"
                    "Puedes hacer clic en 'Exportar Resultados' para guardar la bitácora en Excel.",
                )

        self.root.after(0, finalize)

    def _on_delete_finished(self, stats: dict):
        def finalize():
            self._reset_controls()
            self.btn_export.config(state=tk.NORMAL)

            if stats.get("aborted"):
                self.status_bar_var.set("⏹ Eliminación detenida por el usuario.")
                messagebox.showinfo("Proceso Detenido", f"El proceso de eliminación fue detenido.\n\nClientes procesados: {stats['processed']}/{stats['total']}")
            else:
                self.status_bar_var.set(f"✓ Finalizado: {stats['deleted']} órdenes de eliminación procesadas.")
                messagebox.showinfo(
                    "Eliminación Completada",
                    f"¡Ciclo de eliminación finalizado!\n\n"
                    f"Total registros en archivo: {stats['total']}\n"
                    f"Órdenes de eliminación enviadas: {stats['deleted']}\n"
                    f"Errores: {stats['failed']}\n\n"
                    "Nota: Los clientes que poseían transacciones vinculadas fueron protegidos por SAP.\n"
                    "Puedes exportar la bitácora a Excel.",
                )

        self.root.after(0, finalize)

    def _reset_controls(self):
        self.btn_start.config(state=tk.NORMAL)
        self.btn_pause.config(state=tk.DISABLED, text="⏸ Pausar")
        self.btn_stop.config(state=tk.DISABLED)

    def toggle_pause(self):
        if self.engine.is_paused():
            self.engine.resume()
            self.btn_pause.config(text="⏸ Pausar")
            self.status_bar_var.set("▶ Proceso reanudado.")
        else:
            self.engine.pause()
            self.btn_pause.config(text="▶ Reanudar")
            self.status_bar_var.set("⏸ Proceso en pausa...")

    def stop_process(self):
        self.engine.stop()
        self.status_bar_var.set("Deteniendo bot...")

    def export_results(self):
        mode = self.current_mode_var.get()
        default_name = "resultado_sincronizacion.xlsx" if mode == "sync" else "resultado_eliminacion_clientes.xlsx"
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel (.xlsx)", "*.xlsx")],
            initialfile=default_name,
            title="Guardar resultados",
        )
        if not save_path:
            return

        try:
            records = []
            for item in self.tree.get_children():
                vals = self.tree.item(item)["values"]
                if mode == "sync":
                    records.append({
                        "#": vals[0],
                        "Codigo_SAP": vals[1],
                        "WBCUSTID": vals[2],
                        "SyncFlag": vals[3],
                        "Estado": vals[4],
                        "Detalle": vals[5],
                    })
                else:
                    records.append({
                        "#": vals[0],
                        "Codigo_SAP": vals[1],
                        "Estado": vals[2],
                        "Detalle": vals[3],
                    })
            out_df = pd.DataFrame(records)
            out_df.to_excel(save_path, index=False)
            messagebox.showinfo("Exportación Exitosa", f"Archivo guardado exitosamente en:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Error al exportar", f"No se pudo guardar el archivo:\n{e}")


def main():
    root = tk.Tk()
    app = BotClientesApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
