"""
bot_engine.py
Motor RPA robusto para actualización de Socios de Negocios / Clientes en SAP Business One (HANA).
Permite mapeo dinámico, calibración de puntos y actualización selectiva de:
  - Cabecera: Código de Cliente, RTN
  - Pestaña General: Teléfono 1, Teléfono móvil, Correo electrónico, Estado Activo
  - Panel UDF: WBCUSTID, SyncFlag
  - Pestaña Direcciones: ID de dirección, Calle/Número, Ciudad, Indicador de impuestos
  - Protocolo de auto-recuperación ante errores al Actualizar (Crear ➔ Descartar ➔ Buscar).

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
import re
import json
import time
import threading
from typing import Optional, Callable, Dict, Any, List, Tuple
import pandas as pd
import pyperclip

import win32api
import win32con
import win32gui

import window_manager


class SAPWindowNotFoundError(Exception):
    pass


class BotEngine:
    CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "coordenadas_clientes.json")

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
            "delay_between_fields": 0.22,
            "delay_after_save": 0.85,
            "custom_window_title": None,
            "auto_recover_on_error": True,

            # Coordenadas calibradas (X, Y)
            # 1. Navegación y Control
            "btn_buscar_coord": (2228, -63),          # Lupa / Binoculares en barra de herramientas de SAP
            "btn_crear_coord": (2266, -63),           # Botón Crear / Añadir en barra de SAP (recuperación)
            "btn_confirmar_crear_coord": (2200, 80),   # Botón confirmar 'Crear nuevo' / Descartar
            "card_code_coord": (2135, 16),            # Campo "Código" de Socio de Negocios
            "btn_actualizar_coord": (1977, 905),       # Botón inferior izquierdo ("Buscar" / "Actualizar")
            "barra_estado_coord": (2070, 938),         # Barra de estado inferior de SAP (clic derecho)
            "menu_copiar_error_coord": (2110, 915),    # Opción 'Copiar' en menú contextual (clic izquierdo)
            "tab_general_coord": (2055, 140),          # Pestaña "General"
            "tab_direcciones_coord": (2415, 140),      # Pestaña "Direcciones"

            # 2. Cabecera y Pestaña General
            "rtn_coord": (2135, 86),                  # Campo "RTN" en la cabecera
            "telefono_coord": (2135, 162),            # Campo "Teléfono 1" en pestaña General
            "movil_coord": (2135, 192),               # Campo "Teléfono móvil" en pestaña General
            "correo_coord": (2135, 222),              # Campo "Correo electrónico" en pestaña General
            "activo_coord": (1951, 817),              # Radio button "Activo" en pestaña General
            "wbcustid_coord": (3870, 878),            # Campo "WBCUSTID" en panel UDF
            "syncflag_coord": (3870, 928),            # Campo "SyncFlag" en panel UDF

            # 3. Pestaña Direcciones
            "definir_nuevo_factura_coord": (1990, 224), # Opción 'Definir nuevo' bajo Destinatario de factura
            "id_direccion_coord": (3460, 204),          # Campo "ID de dirección"
            "calle_numero_coord": (3460, 260),          # Campo "Calle/ Número"
            "ciudad_coord": (3460, 289),                # Campo "Ciudad"
            "btn_copiar_direccion_coord": (3520, 909),  # Botón copiar (>>) a Destino
            "indicador_impuestos_coord": (3460, 369),   # Campo "Indicador de impuestos" (en Destino)
        }

        # Cargar coordenadas previas si existen en JSON
        self.load_coordinates_from_file()

        # Estado del mapeo de campos
        self.dataframe: Optional[pd.DataFrame] = None
        self.column_mapping: Dict[str, str] = {
            "card_code": "",
            "rtn": "",
            "telefono": "",
            "movil": "",
            "correo": "",
            "activo": "",
            "wbcustid": "",
            "syncflag": "",
            "id_direccion": "",
            "calle_numero": "",
            "ciudad": "",
            "indicador_impuestos": "",
        }

        # Banderas de activación de cada campo
        self.update_flags: Dict[str, bool] = {
            "rtn": True,
            "telefono": True,
            "movil": True,
            "correo": True,
            "activo": True,
            "wbcustid": True,
            "syncflag": True,
            "id_direccion": True,
            "calle_numero": True,
            "ciudad": True,
            "indicador_impuestos": True,
        }

        # Callback para notificación externa de cambio de pausa (F8, Pausa o botón)
        self.on_pause_callback: Optional[Callable[[bool], None]] = None

    def save_coordinates_to_file(self) -> bool:
        """Guarda las coordenadas actuales en el archivo JSON persistente."""
        try:
            coords = {
                "btn_buscar_coord": self.config.get("btn_buscar_coord"),
                "btn_crear_coord": self.config.get("btn_crear_coord"),
                "btn_confirmar_crear_coord": self.config.get("btn_confirmar_crear_coord"),
                "card_code_coord": self.config.get("card_code_coord"),
                "btn_actualizar_coord": self.config.get("btn_actualizar_coord"),
                "barra_estado_coord": self.config.get("barra_estado_coord"),
                "menu_copiar_error_coord": self.config.get("menu_copiar_error_coord"),
                "tab_general_coord": self.config.get("tab_general_coord"),
                "tab_direcciones_coord": self.config.get("tab_direcciones_coord"),
                "rtn_coord": self.config.get("rtn_coord"),
                "telefono_coord": self.config.get("telefono_coord"),
                "movil_coord": self.config.get("movil_coord"),
                "correo_coord": self.config.get("correo_coord"),
                "activo_coord": self.config.get("activo_coord"),
                "wbcustid_coord": self.config.get("wbcustid_coord"),
                "syncflag_coord": self.config.get("syncflag_coord"),
                "definir_nuevo_factura_coord": self.config.get("definir_nuevo_factura_coord"),
                "id_direccion_coord": self.config.get("id_direccion_coord"),
                "calle_numero_coord": self.config.get("calle_numero_coord"),
                "ciudad_coord": self.config.get("ciudad_coord"),
                "btn_copiar_direccion_coord": self.config.get("btn_copiar_direccion_coord"),
                "indicador_impuestos_coord": self.config.get("indicador_impuestos_coord"),
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
        Carga un archivo Excel o CSV y auto-detecta las mejores coincidencias de columnas para Clientes.
        """
        if file_path.lower().endswith(".csv"):
            df = pd.read_csv(file_path, dtype=str)
        else:
            df = pd.read_excel(file_path, dtype=str)

        df.columns = [str(c).strip() for c in df.columns]

        candidates = {
            "card_code": [
                "codigo_sn", "codigo sn", "cod_sn", "cod sn", "codigosn", "cardcode",
                "card_code", "codigo_cliente", "codigo cliente", "codigo", "código",
                "cliente", "id_cliente", "id_sap", "sap_code"
            ],
            "rtn": [
                "rtn", "id_fiscal", "identificacion", "rfc", "nit", "tax_id", "cai",
                "numero_rtn", "rtn_cliente"
            ],
            "telefono": [
                "telefono_1", "telefono 1", "telefono1", "telefono", "teléfono",
                "tel 1", "tel1", "tel", "phone"
            ],
            "movil": [
                "telefono_movil", "telefono movil", "celular", "movil", "móvil",
                "mobile", "tel_movil", "tel movil", "telefono_2", "telefono 2"
            ],
            "correo": [
                "correo_electronico", "correo electronico", "correo", "email",
                "e-mail", "mail", "correo_cobro", "correo_cliente"
            ],
            "activo": [
                "activo", "estado", "active", "status", "es_activo", "habilitado"
            ],
            "wbcustid": [
                "wbcustid", "u_wbcustid", "id_zoho", "id zoho", "zoho_id", "zoho id",
                "zoho", "record_id", "id_cliente_zoho", "id_interno"
            ],
            "syncflag": [
                "syncflag", "u_syncflag", "sync_flag", "sync", "sincronizado",
                "flag_sincronizacion"
            ],
            "id_direccion": [
                "id_direccion", "id direccion", "direccion_id", "nombre_direccion",
                "tipo_direccion", "address_id", "codigo_direccion"
            ],
            "calle_numero": [
                "calle_numero", "calle numero", "calle", "calle/numero", "calle / numero",
                "direccion", "dirección", "address", "linea_direccion", "dir"
            ],
            "ciudad": [
                "ciudad", "city", "municipio", "poblacion", "localidad", "distrito"
            ],
            "indicador_impuestos": [
                "indicador_impuestos", "indicador impuestos", "indicador_impuesto",
                "impuesto", "tax", "tax_code", "indicador_de_impuestos", "isv",
                "cod_impuesto", "codigo_impuesto"
            ],
        }

        detected_map: Dict[str, str] = {}
        for key, cand_list in candidates.items():
            matched = self._find_col(df.columns, cand_list)
            detected_map[key] = matched or ""

        # Si no detectó card_code, tomar la primera columna
        if not detected_map["card_code"]:
            detected_map["card_code"] = df.columns[0]

        self.dataframe = df
        self.column_mapping = detected_map

        # Actualizar banderas por defecto según las columnas detectadas
        for k in self.update_flags:
            self.update_flags[k] = bool(detected_map.get(k))

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
        """Detiene la ejecución del bot de inmediato."""
        self._stop_event.set()
        self._pause_event.set()
        self.is_running = False

    def pause(self):
        """Pausa temporalmente el proceso."""
        self._pause_event.clear()

    def resume(self):
        """Reanuda el proceso pausado."""
        self._pause_event.set()

    def is_paused(self) -> bool:
        return not self._pause_event.is_set()

    def _clean_field_value(self, key: str, val: Any) -> str:
        """
        Modela y limpia los valores para evitar errores de guardado en SAP Business One.
        SAP B1 rechaza campos numéricos o códigos con decimales (ej. '.0' proveniente de Excel).
        - Para teléfono, móvil, WBCUSTID, RTN, ID dirección, código: elimina decimales conservando ceros a la izquierda.
        - Para campo 'activo': normaliza a 'Y' o 'N'.
        - Para texto libre (correo, calle, ciudad, etc.): elimina decimales residuales si son puramente numéricos terminados en .0.
        """
        if pd.isna(val) or val is None:
            return ""
        s = str(val).strip()
        if not s or s.lower() == "nan":
            return ""

        # Manejo para estado Activo
        if key == "activo":
            if s.upper() in ["Y", "S", "SI", "1", "1.0", "TRUE", "T", "ACTIVO"]:
                return "Y"
            elif s.upper() in ["N", "NO", "0", "0.0", "FALSE", "F", "INACTIVO"]:
                return "N"
            return s

        # Si termina en .0, .00, etc. (típico de lectura de enteros en Excel)
        if re.match(r"^-?\d+\.0+$", s):
            return s.split(".")[0]

        # Manejo para campos de código o identificación donde los ceros a la izquierda son vitales (RTN, card_code, id_direccion)
        if key in ["rtn", "card_code", "id_direccion"]:
            if "." in s:
                parts = s.split(".")
                if parts[1].isdigit():
                    return parts[0]
            return s

        # Manejo para números telefónicos y UDFs numéricos
        if key in ["telefono", "movil", "wbcustid"]:
            if "." in s:
                parts = s.split(".")
                if len(parts) == 2 and parts[1].isdigit():
                    try:
                        f = float(s)
                        if s.startswith("0") and not s.startswith("0."):
                            return parts[0]
                        return str(int(round(f)))
                    except Exception:
                        return parts[0]
            return s

        return s

    def _safe_click(self, coord: Optional[Tuple[int, int]], clicks: int = 1):
        """Ejecuta clics virtuales precisos en Windows sin importar monitor ni posición negativa."""
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

    def _safe_right_click(self, coord: Optional[Tuple[int, int]]):
        """Ejecuta un clic derecho virtual preciso en Windows."""
        if not coord:
            return
        x, y = int(coord[0]), int(coord[1])
        win32api.SetCursorPos((x, y))
        time.sleep(0.08)
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
        time.sleep(0.05)
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
        time.sleep(0.05)

    def _copy_status_bar_message(self) -> str:
        """
        Copia el mensaje de la barra de estado de SAP Business One:
        1. Clic derecho en la barra de estado.
        2. Clic izquierdo en la opción 'Copiar' del menú contextual.
        3. Retorna el texto del portapapeles.
        """
        barra_coord = self.config.get("barra_estado_coord")
        menu_coord = self.config.get("menu_copiar_error_coord")

        if not barra_coord:
            return ""

        try:
            # Limpiar portapapeles para evitar lecturas de valores anteriores
            pyperclip.copy("")
            time.sleep(0.05)

            # 1. Clic derecho en la barra de estado
            self._safe_right_click(barra_coord)
            time.sleep(0.25)

            # 2. Clic izquierdo en la opción Copiar
            if menu_coord:
                self._safe_click(menu_coord)
                time.sleep(0.15)
            else:
                # Si no calibró la opción exacta, presionar 'C' o Enter en el menú contextual
                self._press_key(ord('C'))
                time.sleep(0.15)

            msg = pyperclip.paste().strip()
            return msg
        except Exception:
            return ""

    def _safe_paste(self, text: str):
        """Copia texto al portapapeles y ejecuta Ctrl+V mediante Win32 puro para evitar pérdidas de foco en RDP."""
        if text is None:
            return
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
        1. NUNCA usar Ctrl+A porque en SAP B1 activa 'Añadir documento / Crear nuevo'.
        2. NUNCA enviar 'Delete / Supr' sin flag extendida porque en RDP el teclado numérico envía puntos '.'
        3. Usar END (extendido) para ir al final, seguido de Shift + HOME (extendido) para resaltar todo el texto.
        4. Al pegar con Ctrl+V encima de la selección, el nuevo texto REEMPLAZA el anterior limpiamente.
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

    def _write_field_value(self, coord: Optional[Tuple[int, int]], val: str, field_name: str = ""):
        """
        Escribe un valor en un campo de SAP Business One de forma 100% limpia y atómica:
        1. Doble clic al campo para posicionarse y activar foco.
        2. Selecciona todo el contenido existente con END + Shift+HOME.
        3. Pega el valor nuevo reemplazando la selección con Ctrl+V.
        """
        if not coord or val is None or str(val).strip() == "":
            return

        # 1. Clic directo al campo
        self._safe_click(coord, clicks=2)
        time.sleep(0.10)

        # 2. Resaltar todo el texto existente
        self._select_all_in_field()
        time.sleep(0.04)

        # 3. Pegar el nuevo valor limpio sobre la selección
        self._safe_paste(str(val).strip())
        time.sleep(0.08)

    def _press_enter(self):
        """Presiona ENTER de forma robusta con Win32 y tiempo de retención para RDP."""
        sc = win32api.MapVirtualKey(win32con.VK_RETURN, 0)
        win32api.keybd_event(win32con.VK_RETURN, sc, 0, 0)
        time.sleep(0.08)
        win32api.keybd_event(win32con.VK_RETURN, sc, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.04)

    def _press_key(self, vk: int):
        """Envía una pulsación de tecla virtual limpia."""
        sc = win32api.MapVirtualKey(vk, 0)
        win32api.keybd_event(vk, sc, 0, 0)
        time.sleep(0.04)
        win32api.keybd_event(vk, sc, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.03)

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

    def _press_ctrl_a(self):
        """Presiona Ctrl+A para invocar 'Añadir / Crear nuevo' en SAP B1."""
        sc_ctrl = win32api.MapVirtualKey(win32con.VK_CONTROL, 0)
        sc_a = win32api.MapVirtualKey(ord('A'), 0)
        win32api.keybd_event(win32con.VK_CONTROL, sc_ctrl, 0, 0)
        time.sleep(0.03)
        win32api.keybd_event(ord('A'), sc_a, 0, 0)
        time.sleep(0.03)
        win32api.keybd_event(ord('A'), sc_a, win32con.KEYEVENTF_KEYUP, 0)
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

    def _start_esc_listener(self):
        """Monitorea globalmente F8, Pausa o ESC para pausar/reanudar el bot sin cerrar SAP."""
        def _esc_loop():
            while self.is_running:
                try:
                    f8_pressed = bool(win32api.GetAsyncKeyState(win32con.VK_F8) & 0x8000)
                    pause_pressed = bool(win32api.GetAsyncKeyState(win32con.VK_PAUSE) & 0x8000)
                    esc_pressed = bool(win32api.GetAsyncKeyState(win32con.VK_ESCAPE) & 0x8000)

                    if f8_pressed or pause_pressed or esc_pressed:
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

    def _detect_sap_popup(self) -> Optional[Tuple[int, str]]:
        """Detecta si hay un cuadro de diálogo emergente o mensaje de error de SAP Business One."""
        found = []
        def enum_cb(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                cls = win32gui.GetClassName(hwnd)
                txt = win32gui.GetWindowText(hwnd).strip()
                if cls == "#32770" or "sap business one" in txt.lower():
                    rect = win32gui.GetWindowRect(hwnd)
                    w = rect[2] - rect[0]
                    h = rect[3] - rect[1]
                    if 40 < w < 850 and 40 < h < 600:
                        found.append((hwnd, txt))
            return True
        try:
            win32gui.EnumWindows(enum_cb, None)
        except Exception:
            pass
        return found[0] if found else None

    def _recover_and_reset_to_search(self):
        """
        Protocolo de auto-recuperación ante errores al Actualizar en SAP Business One:
        1. Cierra posibles popups de error emergentes con Enter o Escape.
        2. Selecciona el botón 'Crear cliente / Añadir' en la barra superior (o Ctrl+A).
        3. Si SAP muestra cuadro de confirmación ('¿Desea descartar los cambios y crear nuevo?'),
           selecciona el botón de confirmación calibrado o presiona Enter para descartar.
        4. Selecciona el botón 'Buscar' (Lupa en barra superior o Ctrl+F).
        5. Con esto la pantalla queda completamente limpia en Modo Buscar para pasar al siguiente cliente.
        """
        btn_crear = self.config.get("btn_crear_coord")
        btn_confirmar = self.config.get("btn_confirmar_crear_coord")
        btn_buscar = self.config.get("btn_buscar_coord")

        # 1. Cerrar posibles popups de alerta solo con ESC (NUNCA enviar Enter porque cerraría la ventana de SAP si el botón es OK)
        self._press_key(win32con.VK_ESCAPE)
        time.sleep(0.20)

        # 2. Clic en 'Crear / Añadir'
        if btn_crear:
            self._safe_click(btn_crear)
        else:
            self._press_ctrl_a()
        time.sleep(0.40)

        # 3. Confirmar 'Crear nuevo' / Descartar modificaciones no guardadas si aparece diálogo
        popup = self._detect_sap_popup()
        if popup:
            if btn_confirmar:
                self._safe_click(btn_confirmar)
                time.sleep(0.25)
            else:
                self._press_enter()
                time.sleep(0.25)

        # 4. Clic en 'Buscar' (Lupa) y envío de Ctrl+F nativo para garantizar Modo Buscar
        if btn_buscar:
            self._safe_click(btn_buscar)
            time.sleep(0.15)
        self._press_ctrl_f()
        time.sleep(0.50)

    def _process_single_client(self, row_data: Dict[str, str]) -> List[str]:
        """
        Ejecuta el ciclo de actualización de un socio de negocio individual con precisión:
        1. Clic en botón Buscar (Lupa barra SAP) y envío de Ctrl+F para asegurar Modo Buscar.
        2. Clic en campo 'Código', selecciona y pega el código de cliente.
        3. Clic en botón inferior 'Buscar / Actualizar' para disparar la búsqueda en SAP.
        4. Actualiza campos de Cabecera y pestaña General (RTN, Teléfono, Móvil, Correo, Activo, WBCUSTID, SyncFlag).
        5. Actualiza campos de pestaña Direcciones (ID de dirección, Calle/Número, Ciudad, Indicador de impuestos).
        6. Clic en botón inferior 'Actualizar' (guardar).
        7. Si ocurre error al actualizar, ejecuta auto-recuperación (Crear ➔ Descartar ➔ Buscar).
        """
        btn_buscar = self.config.get("btn_buscar_coord")
        card_coord = self.config.get("card_code_coord")
        btn_actualizar = self.config.get("btn_actualizar_coord")
        tab_general = self.config.get("tab_general_coord")
        tab_dir = self.config.get("tab_direcciones_coord")

        rtn_coord = self.config.get("rtn_coord")
        tel_coord = self.config.get("telefono_coord")
        movil_coord = self.config.get("movil_coord")
        correo_coord = self.config.get("correo_coord")
        activo_coord = self.config.get("activo_coord")
        wb_coord = self.config.get("wbcustid_coord")
        sync_coord = self.config.get("syncflag_coord")

        id_dir_coord = self.config.get("id_direccion_coord")
        calle_coord = self.config.get("calle_numero_coord")
        ciudad_coord = self.config.get("ciudad_coord")
        imp_coord = self.config.get("indicador_impuestos_coord")

        # Limpiar y modelar valores sin decimales no deseados
        card_code = self._clean_field_value("card_code", row_data.get("card_code", ""))
        rtn_val = self._clean_field_value("rtn", row_data.get("rtn", ""))
        tel_val = self._clean_field_value("telefono", row_data.get("telefono", ""))
        movil_val = self._clean_field_value("movil", row_data.get("movil", ""))
        correo_val = self._clean_field_value("correo", row_data.get("correo", ""))
        activo_val = self._clean_field_value("activo", row_data.get("activo", ""))
        wb_val = self._clean_field_value("wbcustid", row_data.get("wbcustid", ""))
        sync_val = self._clean_field_value("syncflag", row_data.get("syncflag", "")) or "T"

        id_dir_val = self._clean_field_value("id_direccion", row_data.get("id_direccion", ""))
        calle_val = self._clean_field_value("calle_numero", row_data.get("calle_numero", ""))
        ciudad_val = self._clean_field_value("ciudad", row_data.get("ciudad", ""))
        imp_val = self._clean_field_value("indicador_impuestos", row_data.get("indicador_impuestos", ""))

        self._check_pause()

        # Si había quedado un popup previo o pantalla sucia, auto-recuperar
        popup_prev = self._detect_sap_popup()
        if popup_prev:
            self._recover_and_reset_to_search()

        # ========================================================
        # PASO 1: Garantizar Modo Buscar en SAP
        # ========================================================
        if btn_buscar:
            self._safe_click(btn_buscar)
            time.sleep(0.15)
        self._press_ctrl_f()
        time.sleep(self.config.get("delay_after_find", 0.45))

        # Verificar si al dar clic en Buscar apareció diálogo de "¿Desea guardar cambios?"
        popup_search = self._detect_sap_popup()
        if popup_search:
            self._recover_and_reset_to_search()

        self._check_pause()

        # ========================================================
        # PASO 2: Clic en campo 'Código de Cliente' y buscar
        # ========================================================
        if card_coord:
            self._safe_click(card_coord, clicks=2)
            time.sleep(0.10)

        # Seleccionar texto previo limpiamente (sin Delete para no escribir puntos)
        self._select_all_in_field()
        time.sleep(0.04)

        # Pegar el código del cliente
        self._safe_paste(card_code)
        time.sleep(0.15)

        # Ejecutar Búsqueda haciendo clic en el botón inferior 'Buscar'
        if btn_actualizar:
            self._safe_click(btn_actualizar)
        else:
            self._press_enter()

        time.sleep(self.config.get("delay_after_search", 1.0))
        self._check_pause()

        actualizados = []

        # ========================================================
        # PASO 3: Cabecera y Pestaña General
        # ========================================================
        # A) RTN en Cabecera
        if self.update_flags.get("rtn") and rtn_val and rtn_coord:
            self._check_pause()
            self._write_field_value(rtn_coord, rtn_val, "RTN")
            actualizados.append(f"RTN={rtn_val}")
            time.sleep(self.config.get("delay_between_fields", 0.22))

        # B) Campos dentro de Pestaña General
        has_general_fields = (
            (self.update_flags.get("telefono") and tel_val and tel_coord) or
            (self.update_flags.get("movil") and movil_val and movil_coord) or
            (self.update_flags.get("correo") and correo_val and correo_coord) or
            (self.update_flags.get("activo") and activo_val and activo_coord)
        )

        if has_general_fields and tab_general:
            self._check_pause()
            self._safe_click(tab_general)
            time.sleep(self.config.get("delay_after_tab_click", 0.35))

        if self.update_flags.get("telefono") and tel_val and tel_coord:
            self._check_pause()
            self._write_field_value(tel_coord, tel_val, "Teléfono")
            actualizados.append(f"Tel={tel_val}")
            time.sleep(self.config.get("delay_between_fields", 0.22))

        if self.update_flags.get("movil") and movil_val and movil_coord:
            self._check_pause()
            self._write_field_value(movil_coord, movil_val, "Móvil")
            actualizados.append(f"Móvil={movil_val}")
            time.sleep(self.config.get("delay_between_fields", 0.22))

        if self.update_flags.get("correo") and correo_val and correo_coord:
            self._check_pause()
            self._write_field_value(correo_coord, correo_val, "Correo")
            actualizados.append(f"Correo={correo_val}")
            time.sleep(self.config.get("delay_between_fields", 0.22))

        if self.update_flags.get("activo") and activo_val and activo_coord:
            self._check_pause()
            if activo_val.upper() in ["Y", "S", "SI", "1", "TRUE", "T", "ACTIVO"]:
                self._safe_click(activo_coord)
                actualizados.append("Activo=Sí")
                time.sleep(self.config.get("delay_between_fields", 0.22))

        # C) Campos de Barra Lateral UDF (WBCUSTID, SyncFlag)
        if self.update_flags.get("wbcustid") and wb_val and wb_coord:
            self._check_pause()
            self._write_field_value(wb_coord, wb_val, "WBCUSTID")
            actualizados.append(f"WBCUSTID={wb_val}")
            time.sleep(self.config.get("delay_between_fields", 0.22))

        if self.update_flags.get("syncflag") and sync_val and sync_coord:
            self._check_pause()
            self._write_field_value(sync_coord, sync_val, "SyncFlag")
            actualizados.append(f"SyncFlag={sync_val}")
            time.sleep(self.config.get("delay_between_fields", 0.22))

        # ========================================================
        # PASO 4: Pestaña Direcciones (Destinatario ➔ Copiar >> ➔ Destino)
        # ========================================================
        has_dir_fields = (
            (self.update_flags.get("id_direccion") and id_dir_val and id_dir_coord) or
            (self.update_flags.get("calle_numero") and calle_val and calle_coord) or
            (self.update_flags.get("ciudad") and ciudad_val and ciudad_coord) or
            (self.update_flags.get("indicador_impuestos") and imp_val and imp_coord)
        )

        if has_dir_fields and tab_dir:
            self._check_pause()
            self._safe_click(tab_dir)
            time.sleep(self.config.get("delay_after_tab_click", 0.40))

            # 4.1 Clic en 'Definir nuevo' bajo Destinatario de Factura si está calibrado
            btn_def_nuevo = self.config.get("definir_nuevo_factura_coord")
            if btn_def_nuevo:
                self._check_pause()
                self._safe_click(btn_def_nuevo)
                time.sleep(0.35)

            # 4.2 Llenar datos de la dirección en Destinatario de Factura (sin impuesto)
            if self.update_flags.get("id_direccion") and id_dir_val and id_dir_coord:
                self._check_pause()
                self._write_field_value(id_dir_coord, id_dir_val, "ID Dirección")
                actualizados.append(f"ID_Dir={id_dir_val}")
                time.sleep(self.config.get("delay_between_fields", 0.22))

            if self.update_flags.get("calle_numero") and calle_val and calle_coord:
                self._check_pause()
                self._write_field_value(calle_coord, calle_val, "Calle/Número")
                actualizados.append(f"Calle={calle_val}")
                time.sleep(self.config.get("delay_between_fields", 0.22))

            if self.update_flags.get("ciudad") and ciudad_val and ciudad_coord:
                self._check_pause()
                self._write_field_value(ciudad_coord, ciudad_val, "Ciudad")
                actualizados.append(f"Ciudad={ciudad_val}")
                time.sleep(self.config.get("delay_between_fields", 0.22))

            # 4.3 Clic en botón Copiar (>>) para replicar a Destino
            btn_copiar = self.config.get("btn_copiar_direccion_coord")
            if btn_copiar:
                self._check_pause()
                self._safe_click(btn_copiar)
                # Esperar a que SAP copie y enfoque automáticamente la dirección en Destino
                time.sleep(0.50)

            # 4.4 En Destino, escribir el Indicador de Impuestos
            if self.update_flags.get("indicador_impuestos") and imp_val and imp_coord:
                self._check_pause()
                self._write_field_value(imp_coord, imp_val, "Impuestos")
                actualizados.append(f"Imp={imp_val}")
                time.sleep(self.config.get("delay_between_fields", 0.22))

        self._check_pause()

        # ========================================================
        # PASO 5: Guardar con "Actualizar"
        # ========================================================
        if btn_actualizar:
            self._safe_click(btn_actualizar)
        else:
            self._press_alt_a()

        time.sleep(self.config.get("delay_after_save", 0.85))

        # ========================================================
        # PASO 6: Captura del Mensaje de SAP
        # ========================================================
        sap_msg = ""
        # Verificar primero si surgió algún cuadro modal de error
        popup = self._detect_sap_popup()
        if popup:
            _, popup_text = popup
            sap_msg = popup_text
            # Cerrar el popup de alerta solo con ESC (sin Enter para no presionar OK en la ventana principal)
            self._press_key(win32con.VK_ESCAPE)
            time.sleep(0.20)

        # Si no hubo popup modal, intentar leer el mensaje de la barra de estado inferior
        if not sap_msg and self.config.get("barra_estado_coord"):
            sap_msg = self._copy_status_bar_message()

        # ========================================================
        # PASO 7: Clasificación de Resultado
        # ========================================================
        msg_lower = sap_msg.lower() if sap_msg else ""
        is_success = any(w in msg_lower for w in ["éxito", "exito", "correcta", "finalizada con éxito", "actualizad"])
        is_error = any(w in msg_lower for w in ["error", "no se puede", "obligatorio", "ya existe", "inválido", "invalido", "rechaz", "falló", "fallo", "bloqueado"])

        # Si detectamos error o vino popup
        if popup or is_error or (sap_msg and not is_success and len(sap_msg) > 6):
            err_desc = sap_msg or (popup[1] if popup else "Error al guardar en SAP Business One")
            # Auto-recuperación SOLO ante error para descartar cambios sucios y regresar a Modo Buscar
            if self.config.get("auto_recover_on_error", True):
                self._recover_and_reset_to_search()
            raise Exception(f"{err_desc}")

        if sap_msg:
            actualizados.append(f"SAP='{sap_msg}'")

        # GUARDADO EXITOSO:
        # NO se toca el botón OK ni se envía Enter para evitar cerrar la ventana de SAP ("sacarlo de pantalla").
        # Al iniciar el siguiente cliente, PASO 1 enviará Ctrl+F nativo para volver a Modo Buscar limpiamente.
        return actualizados

    def run_single_test_client(
        self,
        client_index: int = 0,
        on_countdown_tick: Optional[Callable[[int], None]] = None,
    ) -> Dict[str, Any]:
        """
        Ejecuta la actualización de UN SOLO socio de negocio para pruebas.
        """
        self._stop_event.clear()
        self._pause_event.set()
        self.is_running = True
        self._start_esc_listener()

        df = self.dataframe
        mapping = self.column_mapping

        if df is None or not mapping.get("card_code") or mapping["card_code"] not in df.columns:
            raise ValueError("No hay un archivo cargado con columna válida de Código de Cliente.")

        valid_rows = []
        for idx, row in df.iterrows():
            raw_code = row[mapping["card_code"]] if pd.notna(row[mapping["card_code"]]) else ""
            code = self._clean_field_value("card_code", raw_code)
            if code and code.lower() != "nan":
                row_data = {"card_code": code}
                for k, col_name in mapping.items():
                    if k != "card_code" and col_name and col_name in df.columns and pd.notna(row[col_name]):
                        row_data[k] = self._clean_field_value(k, row[col_name])
                    elif k != "card_code":
                        row_data[k] = ""
                valid_rows.append((idx, row_data))

        if not valid_rows:
            raise ValueError("No se encontraron registros de clientes válidos en el archivo.")

        target_row = valid_rows[min(client_index, len(valid_rows) - 1)]
        row_idx, row_data = target_row

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

        actualizados = self._process_single_client(row_data)
        self.is_running = False

        return {
            "row_idx": row_idx,
            "card_code": row_data["card_code"],
            "row_data": row_data,
            "actualizados": actualizados,
        }

    def run_update_clients(
        self,
        progress_callback: Optional[Callable[[Dict[str, Any]], None]] = None,
        on_countdown_tick: Optional[Callable[[int], None]] = None,
        completion_callback: Optional[Callable[[Dict[str, Any]], None]] = None,
    ):
        """
        Ejecuta el ciclo de actualización masivo para todos los clientes respetando
        únicamente los campos mapeados y activados por el usuario.
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
                    "1. Que el Escritorio Remoto esté abierto y conectado.\n"
                    "2. Que hayas iniciado sesión en SAP Business One.\n"
                    "3. Que la ventana 'Datos maestros socio de negocios' esté abierta."
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
            mapping = self.column_mapping

            if df is None or not mapping.get("card_code") or mapping["card_code"] not in df.columns:
                raise ValueError("El archivo cargado no contiene una columna válida de Código de Cliente.")

            # Filtrar filas válidas
            valid_rows = []
            for idx, row in df.iterrows():
                raw_code = row[mapping["card_code"]] if pd.notna(row[mapping["card_code"]]) else ""
                code = self._clean_field_value("card_code", raw_code)
                if code and code.lower() != "nan":
                    row_data = {"card_code": code}
                    for k, col_name in mapping.items():
                        if k != "card_code" and col_name and col_name in df.columns and pd.notna(row[col_name]):
                            row_data[k] = self._clean_field_value(k, row[col_name])
                        elif k != "card_code":
                            row_data[k] = ""
                    valid_rows.append((idx, row_data))

            stats["total"] = len(valid_rows)

            for item_num, (row_idx, row_data) in enumerate(valid_rows, start=1):
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

                card_code = row_data["card_code"]

                if progress_callback:
                    progress_callback({
                        "index": item_num,
                        "total": len(valid_rows),
                        "row_idx": row_idx,
                        "card_code": card_code,
                        "row_data": row_data,
                        "status": "PROCESANDO",
                        "message": f"Buscando cliente {card_code} en SAP...",
                    })

                try:
                    actualizados = self._process_single_client(row_data)

                    stats["processed"] += 1
                    stats["success"] += 1

                    detalle_str = ", ".join(actualizados) if actualizados else "Sin cambios"
                    if progress_callback:
                        progress_callback({
                            "index": item_num,
                            "total": len(valid_rows),
                            "row_idx": row_idx,
                            "card_code": card_code,
                            "row_data": row_data,
                            "status": "OK",
                            "message": f"Actualizado exitosamente ({detalle_str})",
                        })

                except Exception as row_err:
                    stats["processed"] += 1
                    stats["failed"] += 1

                    # Si ocurrió un error en la fila y auto_recover está activo, limpiar pantalla
                    if self.config.get("auto_recover_on_error", True):
                        try:
                            self._recover_and_reset_to_search()
                        except Exception:
                            pass

                    if progress_callback:
                        progress_callback({
                            "index": item_num,
                            "total": len(valid_rows),
                            "row_idx": row_idx,
                            "card_code": card_code,
                            "row_data": row_data,
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
