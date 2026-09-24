"""
bot_articulos_engine.py
Motor RPA para actualización de parámetros de planificación en Maestro de Artículos (OITM) de SAP Business One (HANA).
Permite mapeo dinámico, calibración de 7 puntos y actualización selectiva de:
  - Botón Buscar (Lupa)
  - Campo Número de artículo (Código)
  - Pestaña Datos de planificación
  - Cantidad de pedido mínimo (Mínimo de compra)
  - Tiempo lead (Lead Time en días)
  - Días de tolerancia (Tiempo de retraso / tolerancia)
  - Botón Actualizar
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
import json
import time
import threading
from typing import Optional, Callable, Dict, Any, List, Tuple
import pandas as pd
import pyautogui
import pyperclip

import win32api
import win32con
import win32gui

import window_manager

pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0.05


class SAPWindowNotFoundError(Exception):
    pass


class BotArticulosEngine:
    CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "coordenadas_articulos.json")

    def __init__(self):
        self._stop_event = threading.Event()
        self._pause_event = threading.Event()
        self._pause_event.set()
        self.is_running = False

        # Configuración de tiempos y coordenadas
        self.config = {
            "countdown_seconds": 3,
            "delay_after_find": 0.45,
            "delay_after_search": 1.0,
            "delay_after_tab_click": 0.4,
            "delay_between_fields": 0.25,
            "delay_after_save": 0.8,
            "custom_window_title": None,
            # Coordenadas calibradas (X, Y)
            "btn_buscar_coord": None,        # 1. Botón Buscar / Lupa en barra de SAP
            "item_code_coord": None,         # 2. Campo "Número de artículo"
            "tab_planificacion_coord": None, # 3. Pestaña "Datos de planificación"
            "min_qty_coord": None,           # 4. Campo "Cantidad pedido mínimo"
            "lead_time_coord": None,         # 5. Campo "Tiempo lead"
            "tolerance_coord": None,         # 6. Campo "Días de tolerancia"
            "btn_actualizar_coord": None,    # 7. Botón "Actualizar"
        }

        # Cargar coordenadas previas si existen
        self.load_coordinates_from_file()

        # Estado del mapeo de campos
        self.dataframe: Optional[pd.DataFrame] = None
        self.item_col: Optional[str] = None
        self.min_qty_col: Optional[str] = None
        self.lead_time_col: Optional[str] = None
        self.tolerance_col: Optional[str] = None

        # Banderas para saber qué campos actualizar en esta ejecución
        self.update_flags = {
            "min_qty": True,
            "lead_time": True,
            "tolerance": True,
        }

        # Callback para notificación externa de cambio de pausa (ej: vía tecla ESC)
        self.on_pause_callback: Optional[Callable[[bool], None]] = None

    def save_coordinates_to_file(self) -> bool:
        """Guarda las coordenadas actuales en el archivo JSON persistente."""
        try:
            coords = {
                "btn_buscar_coord": self.config.get("btn_buscar_coord"),
                "item_code_coord": self.config.get("item_code_coord"),
                "tab_planificacion_coord": self.config.get("tab_planificacion_coord"),
                "min_qty_coord": self.config.get("min_qty_coord"),
                "lead_time_coord": self.config.get("lead_time_coord"),
                "tolerance_coord": self.config.get("tolerance_coord"),
                "btn_actualizar_coord": self.config.get("btn_actualizar_coord"),
            }
            with open(self.CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(coords, f, indent=4)
            return True
        except Exception:
            return False

    def load_coordinates_from_file(self) -> Dict[str, Any]:
        """Carga las coordenadas almacenadas previamente en el archivo JSON."""
        if os.path.exists(self.CONFIG_FILE):
            try:
                with open(self.CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                for k, v in data.items():
                    if k in self.config and v is not None:
                        self.config[k] = tuple(v) if isinstance(v, (list, tuple)) else None
                return data
            except Exception:
                pass
        return {}

    def load_file(self, file_path: str) -> Tuple[pd.DataFrame, Dict[str, str]]:
        """
        Carga un archivo Excel o CSV y auto-detecta las mejores coincidencias de columnas.
        """
        if file_path.lower().endswith(".csv"):
            df = pd.read_csv(file_path, dtype=str)
        else:
            df = pd.read_excel(file_path, dtype=str)

        df.columns = [str(c).strip() for c in df.columns]

        # 1. Código de Artículo
        item_candidates = [
            "codigo_articulo", "codigo articulo", "codigo_item", "itemcode",
            "item_code", "articulo", "artículo", "codigo", "código", "item",
            "sku", "id_articulo", "item_id"
        ]
        detected_item = self._find_col(df.columns, item_candidates) or df.columns[0]

        # 2. Mínimo de Compra
        min_candidates = [
            "minimo_compra", "minimo compra", "pedido_minimo", "pedido minimo",
            "min_order_qty", "min_qty", "cantidad_minima", "cantidad minima",
            "cant_minima", "minorderqty", "minimo", "mínimo", "moq", "pedido_min"
        ]
        detected_min = self._find_col(df.columns, min_candidates)

        # 3. Tiempo Lead
        lead_candidates = [
            "tiempo_lead", "tiempo lead", "lead_time", "leadtime", "lead time",
            "lead", "dias_lead", "dias lead", "dias_entrega", "dias de entrega",
            "tiempo_entrega", "plazo_entrega", "lt prom", "lt_prom", "lt", "lead time prom"
        ]
        detected_lead = self._find_col(df.columns, lead_candidates)

        # 4. Días de Tolerancia
        tol_candidates = [
            "dias_tolerancia", "dias tolerancia", "tolerancia", "tolerance_days",
            "tolerance", "dias de tolerancia", "tiempo_retraso", "retraso",
            "dias_retraso", "dias de retraso", "tiempo_tolerancia", "retraso mas alto", "retraso_mas_alto"
        ]
        detected_tol = self._find_col(df.columns, tol_candidates)

        self.dataframe = df
        self.item_col = detected_item
        self.min_qty_col = detected_min
        self.lead_time_col = detected_lead
        self.tolerance_col = detected_tol

        # Actualizar banderas por defecto según detección
        self.update_flags["min_qty"] = bool(detected_min)
        self.update_flags["lead_time"] = bool(detected_lead)
        self.update_flags["tolerance"] = bool(detected_tol)

        detected_map = {
            "item_col": detected_item,
            "min_qty_col": detected_min or "",
            "lead_time_col": detected_lead or "",
            "tolerance_col": detected_tol or "",
        }
        return df, detected_map

    def _find_col(self, columns: List[str], candidates: List[str]) -> Optional[str]:
        for cand in candidates:
            for col in columns:
                if cand == col.lower().strip():
                    return col
        for cand in candidates:
            for col in columns:
                if cand in col.lower().strip():
                    return col
        return None

    def stop(self):
        self._stop_event.set()
        self._pause_event.set()
        self.is_running = False

    def pause(self):
        self._pause_event.clear()

    def resume(self):
        self._pause_event.set()

    def is_paused(self) -> bool:
        return not self._pause_event.is_set()

    def _safe_paste(self, text: str):
        """Copia texto al portapapeles y ejecuta Ctrl+V mediante Win32 puro para evitar pérdida de foco en RDP."""
        pyperclip.copy(str(text))
        time.sleep(0.04)
        sc_ctrl = win32api.MapVirtualKey(win32con.VK_CONTROL, 0)
        sc_v = win32api.MapVirtualKey(ord('V'), 0)
        win32api.keybd_event(win32con.VK_CONTROL, sc_ctrl, 0, 0)
        time.sleep(0.02)
        win32api.keybd_event(ord('V'), sc_v, 0, 0)
        time.sleep(0.02)
        win32api.keybd_event(ord('V'), sc_v, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.02)
        win32api.keybd_event(win32con.VK_CONTROL, sc_ctrl, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.05)

    def _select_all_in_field(self):
        """
        Selecciona todo el texto existente en el campo activo de SAP B1 de forma 100% segura.
        Reglas críticas para SAP B1 en RDP:
        1. NUNCA usar Ctrl+A porque en SAP B1 es el atajo de 'Añadir documento / Crear artículo'.
        2. NUNCA enviar 'Delete / Supr' sin flag extendida porque en RDP el teclado numérico envía puntos '.'
        3. Usar END (extendido) para ir al final, seguido de Shift + HOME (extendido) para resaltar todo el texto.
        4. Al pegar con Ctrl+V encima de la selección, el nuevo texto REEMPLAZA el anterior limpiamente sin generar puntos ni máscaras.
        """
        sc_end = win32api.MapVirtualKey(win32con.VK_END, 0)
        sc_home = win32api.MapVirtualKey(win32con.VK_HOME, 0)
        sc_shift = win32api.MapVirtualKey(win32con.VK_SHIFT, 0)

        # 1. Enviar END extendido para posicionarse al final del texto actual
        win32api.keybd_event(win32con.VK_END, sc_end, win32con.KEYEVENTF_EXTENDEDKEY, 0)
        time.sleep(0.02)
        win32api.keybd_event(win32con.VK_END, sc_end, win32con.KEYEVENTF_EXTENDEDKEY | win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.03)

        # 2. Shift + HOME extendido para seleccionar todo hacia el inicio
        win32api.keybd_event(win32con.VK_SHIFT, sc_shift, 0, 0)
        time.sleep(0.02)
        win32api.keybd_event(win32con.VK_HOME, sc_home, win32con.KEYEVENTF_EXTENDEDKEY, 0)
        time.sleep(0.02)
        win32api.keybd_event(win32con.VK_HOME, sc_home, win32con.KEYEVENTF_EXTENDEDKEY | win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.02)
        win32api.keybd_event(win32con.VK_SHIFT, sc_shift, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.04)

    def _press_enter(self):
        """Presiona ENTER de forma robusta con Win32."""
        sc = win32api.MapVirtualKey(win32con.VK_RETURN, 0)
        win32api.keybd_event(win32con.VK_RETURN, sc, 0, 0)
        time.sleep(0.03)
        win32api.keybd_event(win32con.VK_RETURN, sc, win32con.KEYEVENTF_KEYUP, 0)

    def _press_ctrl_f(self):
        """Presiona Ctrl+F para entrar a Modo Buscar en SAP B1."""
        sc_ctrl = win32api.MapVirtualKey(win32con.VK_CONTROL, 0)
        sc_f = win32api.MapVirtualKey(ord('F'), 0)
        win32api.keybd_event(win32con.VK_CONTROL, sc_ctrl, 0, 0)
        time.sleep(0.03)
        win32api.keybd_event(ord('F'), sc_f, 0, 0)
        time.sleep(0.03)
        win32api.keybd_event(ord('F'), sc_f, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.02)
        win32api.keybd_event(win32con.VK_CONTROL, sc_ctrl, win32con.KEYEVENTF_KEYUP, 0)

    def _press_alt_a(self):
        """Presiona Alt+A para Guardar/Actualizar en SAP B1."""
        sc_menu = win32api.MapVirtualKey(win32con.VK_MENU, 0)
        sc_a = win32api.MapVirtualKey(ord('A'), 0)
        win32api.keybd_event(win32con.VK_MENU, sc_menu, 0, 0)
        time.sleep(0.03)
        win32api.keybd_event(ord('A'), sc_a, 0, 0)
        time.sleep(0.03)
        win32api.keybd_event(ord('A'), sc_a, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.02)
        win32api.keybd_event(win32con.VK_MENU, sc_menu, win32con.KEYEVENTF_KEYUP, 0)

    def _write_field_value(self, coord: Optional[Tuple[int, int]], val: str, field_name: str = ""):
        """
        Escribe un valor en un campo de SAP Business One de forma 100% limpia y atómica:
        1. Doble clic al campo para posicionarse y activar foco.
        2. Selecciona todo el contenido existente con END + Shift+HOME.
        3. Pega el valor nuevo reemplazando la selección con Ctrl+V.
        """
        if not coord or not val:
            return

        # 1. Clic directo al campo
        self._safe_click(coord, clicks=2)
        time.sleep(0.10)

        # 2. Resaltar todo el texto existente
        self._select_all_in_field()
        time.sleep(0.04)

        # 3. Pegar el nuevo valor limpio sobre la selección
        self._safe_paste(str(val))
        time.sleep(0.08)

    def _clean_numeric_val(self, val: Any) -> str:
        """
        Limpia un valor numérico proveniente de Excel / CSV.
        Convierte valores como '100.0', '100.00', '100', 100.0 en '100' limpio,
        sin decimales residuales ni puntos flotantes no deseados.
        """
        if pd.isna(val) or val is None:
            return ""
        s = str(val).strip()
        if not s or s.lower() == "nan":
            return ""
        try:
            cleaned_s = s.replace(",", ".")
            f = float(cleaned_s)
            if f.is_integer():
                return str(int(f))
            return f"{f:g}"
        except Exception:
            if s.endswith(".0"):
                s = s[:-2]
            return s

    def _safe_click(self, coord: Optional[Tuple[int, int]], clicks: int = 1):
        if not coord:
            return
        x, y = int(coord[0]), int(coord[1])
        win32api.SetCursorPos((x, y))
        time.sleep(0.08)
        for _ in range(clicks):
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            time.sleep(0.05)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            if clicks > 1:
                time.sleep(0.08)

    def _start_esc_listener(self):
        """Monitorea globalmente F8, Pausa o ESC para pausar/reanudar el bot sin cerrar SAP."""
        def _esc_loop():
            while self.is_running:
                try:
                    f8_pressed = bool(win32api.GetAsyncKeyState(win32con.VK_F8) & 0x8000)
                    pause_pressed = bool(win32api.GetAsyncKeyState(win32con.VK_PAUSE) & 0x8000)
                    esc_pressed = bool(win32api.GetAsyncKeyState(win32con.VK_ESCAPE) & 0x8000)

                    if f8_pressed or pause_pressed or esc_pressed:
                        # Debounce para evitar registrar múltiples pulsaciones rápidas
                        time.sleep(0.35)
                        if self.is_paused():
                            self.resume()
                            if self.on_pause_callback:
                                self.on_pause_callback(False)
                        else:
                            self.pause()
                            if self.on_pause_callback:
                                self.on_pause_callback(True)
                except Exception:
                    pass
                time.sleep(0.05)

        t = threading.Thread(target=_esc_loop, daemon=True)
        t.start()

    def _check_pause(self):
        """Pausa la ejecución si se activó la pausa con F8, ESC o botón."""
        while not self._pause_event.is_set():
            time.sleep(0.15)
            if self._stop_event.is_set():
                break

    def _process_single_item(self, item_code: str, min_val: str, lead_val: str, tol_val: str) -> List[str]:
        """
        Ejecuta el ciclo de actualización de un artículo individual con precisión:
        1. Clic en botón Buscar (Lupa) para garantizar Modo Buscar.
        2. Clic en campo 'Número de artículo', selecciona y pega el código sin generar puntos.
        3. Presiona Enter para buscar.
        4. Clic en pestaña 'Datos de planificación'.
        5. Pega Mínimo de compra directamente sobre selección (si está activo).
        6. Pega Tiempo lead directamente sobre selección (si está activo).
        7. Pega Días de tolerancia directamente sobre selección (si está activo).
        8. Guarda con botón 'Actualizar' (o Alt+A).
        """
        btn_buscar = self.config.get("btn_buscar_coord")
        item_coord = self.config.get("item_code_coord")
        tab_coord = self.config.get("tab_planificacion_coord")
        min_coord = self.config.get("min_qty_coord")
        lead_coord = self.config.get("lead_time_coord")
        tol_coord = self.config.get("tolerance_coord")
        btn_actualizar = self.config.get("btn_actualizar_coord")

        self._check_pause()

        # ========================================================
        # PASO 1: Garantizar Modo Buscar en SAP
        # ========================================================
        if btn_buscar:
            self._safe_click(btn_buscar)
        else:
            self._press_ctrl_f()
        time.sleep(self.config.get("delay_after_find", 0.45))

        self._check_pause()

        # ========================================================
        # PASO 2: Clic directo en el Campo "Número de artículo"
        # ========================================================
        if item_coord:
            self._safe_click(item_coord, clicks=2)
            time.sleep(0.10)

        # Seleccionar texto previo limpiamente (sin Delete para no escribir puntos)
        self._select_all_in_field()
        time.sleep(0.04)

        # Pegar el código de artículo y buscar
        self._safe_paste(item_code)
        time.sleep(0.12)
        self._press_enter()
        time.sleep(self.config.get("delay_after_search", 1.0))

        self._check_pause()

        # ========================================================
        # PASO 3: Clic en pestaña "Datos de planificación"
        # ========================================================
        if tab_coord:
            self._safe_click(tab_coord)
            time.sleep(self.config.get("delay_after_tab_click", 0.4))

        actualizados = []

        # ========================================================
        # PASO 4: Campo "Cantidad de pedido mínimo" (si está activo)
        # ========================================================
        if self.update_flags.get("min_qty") and min_val and min_coord:
            self._check_pause()
            self._write_field_value(min_coord, min_val, "Cantidad mínima")
            actualizados.append(f"Mín={min_val}")
            time.sleep(self.config.get("delay_between_fields", 0.25))

        # ========================================================
        # PASO 5: Campo "Tiempo lead" (si está activo)
        # ========================================================
        if self.update_flags.get("lead_time") and lead_val and lead_coord:
            self._check_pause()
            self._write_field_value(lead_coord, lead_val, "Tiempo lead")
            actualizados.append(f"Lead={lead_val}")
            time.sleep(self.config.get("delay_between_fields", 0.25))

        # ========================================================
        # PASO 6: Campo "Días de tolerancia" (si está activo)
        # ========================================================
        if self.update_flags.get("tolerance") and tol_val and tol_coord:
            self._check_pause()
            self._write_field_value(tol_coord, tol_val, "Días tolerancia")
            actualizados.append(f"Tol={tol_val}")
            time.sleep(self.config.get("delay_between_fields", 0.25))

        self._check_pause()

        # ========================================================
        # PASO 7: Guardar con "Actualizar"
        # ========================================================
        if btn_actualizar:
            self._safe_click(btn_actualizar)
        else:
            self._press_alt_a()
        time.sleep(self.config.get("delay_after_save", 0.8))

        return actualizados

    def run_single_test_item(
        self,
        item_index: int = 0,
        on_countdown_tick: Optional[Callable[[int], None]] = None,
    ) -> Dict[str, Any]:
        """
        Ejecuta la actualización de UN SOLO artículo para pruebas.
        """
        self._stop_event.clear()
        self._pause_event.set()
        self.is_running = True
        self._start_esc_listener()

        df = self.dataframe
        item_col = self.item_col
        min_col = self.min_qty_col if self.update_flags.get("min_qty") else None
        lead_col = self.lead_time_col if self.update_flags.get("lead_time") else None
        tol_col = self.tolerance_col if self.update_flags.get("tolerance") else None

        if df is None or not item_col or item_col not in df.columns:
            raise ValueError("No hay un archivo cargado con columna válida de Código de Artículo.")

        # Obtener el artículo seleccionado o el primero válido
        valid_rows = []
        for idx, row in df.iterrows():
            item_code = str(row[item_col]).strip() if pd.notna(row[item_col]) else ""
            if item_code and item_code.lower() != "nan":
                min_val = self._clean_numeric_val(row[min_col]) if min_col and min_col in df.columns else ""
                lead_val = self._clean_numeric_val(row[lead_col]) if lead_col and lead_col in df.columns else ""
                tol_val = self._clean_numeric_val(row[tol_col]) if tol_col and tol_col in df.columns else ""
                valid_rows.append((idx, item_code, min_val, lead_val, tol_val))

        if not valid_rows:
            raise ValueError("No se encontraron registros de artículos válidos en el archivo.")

        target_row = valid_rows[min(item_index, len(valid_rows) - 1)]
        row_idx, item_code, min_val, lead_val, tol_val = target_row

        # Enfocar ventana SAP
        target_window = window_manager.find_target_window(self.config.get("custom_window_title"))
        if not target_window:
            raise SAPWindowNotFoundError(
                "No se detectó la ventana de SAP Business One ni la sesión de Escritorio Remoto abierta."
            )
        hwnd, _ = target_window
        window_manager.bring_window_to_front(hwnd)

        # Cuenta regresiva
        countdown = self.config.get("countdown_seconds", 3)
        for sec in range(countdown, 0, -1):
            if on_countdown_tick:
                on_countdown_tick(sec)
            time.sleep(1.0)
        if on_countdown_tick:
            on_countdown_tick(0)

        window_manager.bring_window_to_front(hwnd)
        time.sleep(0.3)

        actualizados = self._process_single_item(item_code, min_val, lead_val, tol_val)
        self.is_running = False

        return {
            "row_idx": row_idx,
            "item_code": item_code,
            "min_val": min_val,
            "lead_val": lead_val,
            "tol_val": tol_val,
            "actualizados": actualizados,
        }

    def run_update_items(
        self,
        progress_callback: Optional[Callable[[Dict[str, Any]], None]] = None,
        on_countdown_tick: Optional[Callable[[int], None]] = None,
        completion_callback: Optional[Callable[[Dict[str, Any]], None]] = None,
    ):
        """
        Ejecuta el ciclo de actualización de parámetros en el Maestro de Artículos
        respetando únicamente los campos mapeados y activados por el usuario.
        """
        self._stop_event.clear()
        self._pause_event.set()
        self.is_running = True
        self._start_esc_listener()

        stats = {
            "total": 0,
            "processed": 0,
            "success": 0,
            "failed": 0,
            "aborted": False,
            "error_msg": None,
        }

        try:
            # 1. Validar ventana de SAP
            target_window = window_manager.find_target_window(self.config.get("custom_window_title"))
            if not target_window:
                raise SAPWindowNotFoundError(
                    "No se detectó la ventana de SAP Business One ni la sesión de Escritorio Remoto abierta.\n\n"
                    "Por favor verifica:\n"
                    "1. Que el Escritorio Remoto (182.160.29.90) esté abierto y conectado.\n"
                    "2. Que hayas iniciado sesión con tu usuario en SAP Business One.\n"
                    "3. Que la ventana 'Datos maestros de artículo' esté visible en tu pantalla."
                )

            hwnd, title = target_window

            if not window_manager.bring_window_to_front(hwnd):
                raise Exception(f"No fue posible enfocar la ventana: '{title}'")

            # 2. Cuenta regresiva para dar foco
            countdown = self.config.get("countdown_seconds", 3)
            for sec in range(countdown, 0, -1):
                if self._stop_event.is_set():
                    stats["aborted"] = True
                    return
                if on_countdown_tick:
                    on_countdown_tick(sec)
                time.sleep(1.0)

            if on_countdown_tick:
                on_countdown_tick(0)

            window_manager.bring_window_to_front(hwnd)
            time.sleep(0.3)

            df = self.dataframe
            item_col = self.item_col
            min_col = self.min_qty_col if self.update_flags.get("min_qty") else None
            lead_col = self.lead_time_col if self.update_flags.get("lead_time") else None
            tol_col = self.tolerance_col if self.update_flags.get("tolerance") else None

            if df is None or not item_col or item_col not in df.columns:
                raise ValueError("El archivo cargado no contiene una columna válida de Código de Artículo.")

            # Filtrar filas válidas
            valid_rows = []
            for idx, row in df.iterrows():
                item_code = str(row[item_col]).strip() if pd.notna(row[item_col]) else ""
                min_val = self._clean_numeric_val(row[min_col]) if min_col and min_col in df.columns else ""
                lead_val = self._clean_numeric_val(row[lead_col]) if lead_col and lead_col in df.columns else ""
                tol_val = self._clean_numeric_val(row[tol_col]) if tol_col and tol_col in df.columns else ""

                if item_code and item_code.lower() != "nan":
                    valid_rows.append((idx, item_code, min_val, lead_val, tol_val))

            stats["total"] = len(valid_rows)

            for item_num, (row_idx, item_code, min_val, lead_val, tol_val) in enumerate(valid_rows, start=1):
                if self._stop_event.is_set():
                    stats["aborted"] = True
                    break

                while not self._pause_event.is_set():
                    time.sleep(0.2)
                    if self._stop_event.is_set():
                        stats["aborted"] = True
                        break
                if stats["aborted"]:
                    break

                if progress_callback:
                    progress_callback({
                        "index": item_num,
                        "total": len(valid_rows),
                        "row_idx": row_idx,
                        "item_code": item_code,
                        "min_val": min_val,
                        "lead_val": lead_val,
                        "tol_val": tol_val,
                        "status": "PROCESANDO",
                        "message": f"Buscando artículo {item_code} en SAP...",
                    })

                try:
                    actualizados = self._process_single_item(item_code, min_val, lead_val, tol_val)

                    stats["processed"] += 1
                    stats["success"] += 1

                    detalle_str = ", ".join(actualizados) if actualizados else "Sin cambios"
                    if progress_callback:
                        progress_callback({
                            "index": item_num,
                            "total": len(valid_rows),
                            "row_idx": row_idx,
                            "item_code": item_code,
                            "min_val": min_val,
                            "lead_val": lead_val,
                            "tol_val": tol_val,
                            "status": "OK",
                            "message": f"Actualizado exitosamente ({detalle_str})",
                        })

                except Exception as row_err:
                    stats["processed"] += 1
                    stats["failed"] += 1
                    if progress_callback:
                        progress_callback({
                            "index": item_num,
                            "total": len(valid_rows),
                            "row_idx": row_idx,
                            "item_code": item_code,
                            "min_val": min_val,
                            "lead_val": lead_val,
                            "tol_val": tol_val,
                            "status": "ERROR",
                            "message": str(row_err),
                        })

        except Exception as e:
            stats["error_msg"] = str(e)
            stats["aborted"] = True
            raise e
        finally:
            self.is_running = False
            if completion_callback:
                completion_callback(stats)
