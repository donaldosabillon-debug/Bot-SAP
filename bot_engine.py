"""
bot_engine.py
Motor de automatización RPA para SAP Business One (HANA).
Soporta:
  1. Modo Sincronización: Digitación de WBCUSTID (Zoho ID) y SyncFlag = 'T'.
  2. Modo Eliminación: Búsqueda y eliminación de Socios de Negocios duplicados.
"""

import time
import threading
from typing import Optional, Callable, Dict, Any, List, Tuple
import pandas as pd
import pyautogui
import pyperclip

import window_manager

# Configurar seguridad básica de PyAutoGUI
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.05


class SAPWindowNotFoundError(Exception):
    """Excepción lanzada cuando no se detecta la ventana de SAP ni Escritorio Remoto."""
    pass


class BotEngine:
    def __init__(self):
        self._stop_event = threading.Event()
        self._pause_event = threading.Event()
        self._pause_event.set()  # Por defecto no pausado
        self.is_running = False

        # Configuración general y de tiempos
        self.config = {
            "countdown_seconds": 3,
            "delay_after_find": 0.4,          # Tiempo tras presionar Ctrl+F
            "delay_after_search": 0.9,        # Tiempo para que SAP cargue el socio
            "delay_after_field_focus": 0.2,   # Tiempo tras enfocar WBCUSTID
            "delay_between_fields": 0.2,      # Tiempo entre WBCUSTID y SyncFlag
            "delay_after_save": 0.7,          # Tiempo tras presionar Actualizar
            "action_save": "alt_a",           # 'alt_a', 'enter', o 'click'
            "zoho_field_coord": None,         # Tupla (x, y) de la posición de WBCUSTID
            "syncflag_mode": "tab",           # 'tab' o 'click'
            "syncflag_tab_count": 1,          # Tabs necesarios para llegar a SyncFlag
            "syncflag_coord": None,           # Tupla (x, y) de SyncFlag si modo es click
            "syncflag_value": "T",            # Valor a asignar (T = True / Sincronizado)
            "update_syncflag": True,          # Si debe actualizar SyncFlag
            "save_btn_coord": None,           # Tupla (x, y) opcional del botón Actualizar
            "custom_window_title": None,      # Título específico opcional
            # Parámetros para Modo Eliminación
            "delete_method": "menu_alt_d",    # 'menu_alt_d' o 'right_click'
            "delete_click_coord": None,       # Tupla (x, y) para clic derecho en cabecera
            "delay_after_delete_action": 0.5, # Espera tras invocar menú eliminar
            "delay_after_confirm_delete": 0.8,# Espera tras confirmar eliminación
        }

        # Datos cargados
        self.dataframe: Optional[pd.DataFrame] = None
        self.sap_col: Optional[str] = None
        self.zoho_col: Optional[str] = None

    def load_file(self, file_path: str) -> Tuple[pd.DataFrame, str, str]:
        """
        Carga un archivo Excel o CSV y auto-detecta las columnas de Código SAP y WBCUSTID/Zoho ID.
        Retorna (df, sap_col_name, zoho_col_name).
        """
        if file_path.lower().endswith(".csv"):
            df = pd.read_csv(file_path, dtype=str)
        else:
            df = pd.read_excel(file_path, dtype=str)

        # Limpiar nombres de columnas
        df.columns = [str(c).strip() for c in df.columns]

        # Auto-detectar columna de Código SAP
        sap_candidates = [
            "codigo_sn", "codigo sn", "cod_sn", "cod sn", "codigosn",
            "codigo_sap", "codigosap", "cardcode", "codigo", "código",
            "cod_cliente", "cod cliente", "cliente", "id_sap", "sap_code", "sap"
        ]
        detected_sap = None
        for cand in sap_candidates:
            for col in df.columns:
                if cand == col.lower().strip():
                    detected_sap = col
                    break
            if detected_sap:
                break
        if not detected_sap:
            for cand in sap_candidates:
                for col in df.columns:
                    if cand in col.lower().strip():
                        detected_sap = col
                        break
                if detected_sap:
                    break

        # Auto-detectar columna de WBCUSTID / ID Zoho
        zoho_candidates = [
            "wbcustid", "u_wbcustid", "id_zoho", "id zoho", "idzoho",
            "zoho_id", "zoho id", "id_interno", "id interno", "record_id", "record id", "zoho", "id"
        ]
        detected_zoho = None
        for cand in zoho_candidates:
            for col in df.columns:
                if cand == col.lower().strip():
                    detected_zoho = col
                    break
            if detected_zoho:
                break
        if not detected_zoho:
            for cand in zoho_candidates:
                for col in df.columns:
                    if cand in col.lower().strip():
                        detected_zoho = col
                        break
                if detected_zoho:
                    break

        if not detected_sap:
            detected_sap = df.columns[0]
        if not detected_zoho:
            detected_zoho = df.columns[1] if len(df.columns) > 1 else df.columns[0]

        self.dataframe = df
        self.sap_col = detected_sap
        self.zoho_col = detected_zoho

        return df, detected_sap, detected_zoho

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

    def _safe_paste(self, text: str):
        """Copia texto al portapapeles y ejecuta Ctrl+V para evitar fallas de tipeo por latencia en RDP."""
        pyperclip.copy(str(text))
        time.sleep(0.04)
        pyautogui.hotkey("ctrl", "v")

    def _prepare_target_window(self, on_countdown_tick: Optional[Callable[[int], None]] = None) -> Tuple[int, str]:
        """Valida que SAP esté abierto y trae la ventana al frente con cuenta regresiva."""
        target_window = window_manager.find_target_window(self.config.get("custom_window_title"))
        if not target_window:
            raise SAPWindowNotFoundError(
                "No se detectó la ventana de SAP Business One ni la sesión de Escritorio Remoto abierta.\n\n"
                "Por favor verifica:\n"
                "1. Que el Escritorio Remoto (182.160.29.90) esté abierto y conectado.\n"
                "2. Que hayas iniciado sesión con tu usuario en SAP Business One.\n"
                "3. Que la ventana 'Datos maestros de socio de negocios' esté visible."
            )

        hwnd, title = target_window

        if not window_manager.bring_window_to_front(hwnd):
            raise Exception(f"No fue posible traer al frente la ventana: '{title}'")

        countdown = self.config.get("countdown_seconds", 3)
        for sec in range(countdown, 0, -1):
            if self._stop_event.is_set():
                raise KeyboardInterrupt("Proceso cancelado por el usuario")
            if on_countdown_tick:
                on_countdown_tick(sec)
            time.sleep(1.0)

        if on_countdown_tick:
            on_countdown_tick(0)

        window_manager.bring_window_to_front(hwnd)
        time.sleep(0.3)
        return hwnd, title

    # =========================================================================
    # MODO 1: SINCRONIZACIÓN ZOHO (WBCUSTID + SyncFlag = 'T')
    # =========================================================================

    def run_sync(
        self,
        progress_callback: Optional[Callable[[Dict[str, Any]], None]] = None,
        on_countdown_tick: Optional[Callable[[int], None]] = None,
        completion_callback: Optional[Callable[[Dict[str, Any]], None]] = None,
    ):
        """Ejecuta el ciclo de sincronización de IDs y banderas."""
        self._stop_event.clear()
        self._pause_event.set()
        self.is_running = True

        stats = {"total": 0, "processed": 0, "success": 0, "failed": 0, "aborted": False, "error_msg": None}

        try:
            hwnd, title = self._prepare_target_window(on_countdown_tick)

            df = self.dataframe
            sap_col = self.sap_col
            zoho_col = self.zoho_col

            if df is None or sap_col not in df.columns or zoho_col not in df.columns:
                raise ValueError("El archivo de datos no ha sido cargado o las columnas no son válidas.")

            valid_rows = []
            for idx, row in df.iterrows():
                sap_val = str(row[sap_col]).strip() if pd.notna(row[sap_col]) else ""
                zoho_val = str(row[zoho_col]).strip() if pd.notna(row[zoho_col]) else ""
                if zoho_val.endswith(".0"):
                    zoho_val = zoho_val[:-2]
                if sap_val and sap_val.lower() != "nan" and zoho_val and zoho_val.lower() != "nan":
                    valid_rows.append((idx, sap_val, zoho_val))

            stats["total"] = len(valid_rows)

            for item_num, (row_idx, card_code, zoho_id) in enumerate(valid_rows, start=1):
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
                        "card_code": card_code,
                        "zoho_id": zoho_id,
                        "syncflag": self.config.get("syncflag_value", "T"),
                        "status": "PROCESANDO",
                        "message": "Buscando cliente en SAP...",
                    })

                try:
                    # 1. Modo Buscar en SAP
                    pyautogui.hotkey("ctrl", "f")
                    time.sleep(self.config.get("delay_after_find", 0.4))

                    # 2. Digitar Código de Cliente + Enter
                    self._safe_paste(card_code)
                    time.sleep(0.1)
                    pyautogui.press("enter")
                    time.sleep(self.config.get("delay_after_search", 0.9))

                    # 3. Posicionarse en WBCUSTID y pegar
                    zoho_coord = self.config.get("zoho_field_coord")
                    if zoho_coord:
                        pyautogui.click(zoho_coord[0], zoho_coord[1])
                        time.sleep(self.config.get("delay_after_field_focus", 0.2))
                    else:
                        pyautogui.press("tab")
                        time.sleep(0.1)

                    pyautogui.hotkey("ctrl", "a")
                    time.sleep(0.04)
                    self._safe_paste(zoho_id)
                    time.sleep(self.config.get("delay_between_fields", 0.2))

                    # 4. Posicionarse en SyncFlag y poner 'T'
                    if self.config.get("update_syncflag", True):
                        sync_mode = self.config.get("syncflag_mode", "tab")
                        sync_coord = self.config.get("syncflag_coord")
                        tab_count = self.config.get("syncflag_tab_count", 1)

                        if sync_mode == "click" and sync_coord:
                            pyautogui.click(sync_coord[0], sync_coord[1])
                            time.sleep(0.15)
                        else:
                            for _ in range(max(1, tab_count)):
                                pyautogui.press("tab")
                                time.sleep(0.06)

                        pyautogui.hotkey("ctrl", "a")
                        time.sleep(0.04)
                        flag_val = str(self.config.get("syncflag_value", "T"))
                        self._safe_paste(flag_val)
                        time.sleep(0.15)

                    # 5. Guardar / Actualizar
                    action_save = self.config.get("action_save", "alt_a")
                    save_coord = self.config.get("save_btn_coord")

                    if action_save == "click" and save_coord:
                        pyautogui.click(save_coord[0], save_coord[1])
                    elif action_save == "enter":
                        pyautogui.press("enter")
                    else:
                        pyautogui.hotkey("alt", "a")

                    time.sleep(self.config.get("delay_after_save", 0.7))

                    stats["processed"] += 1
                    stats["success"] += 1

                    if progress_callback:
                        progress_callback({
                            "index": item_num,
                            "total": len(valid_rows),
                            "row_idx": row_idx,
                            "card_code": card_code,
                            "zoho_id": zoho_id,
                            "syncflag": self.config.get("syncflag_value", "T"),
                            "status": "OK",
                            "message": f"WBCUSTID={zoho_id} | SyncFlag='T' guardados",
                        })

                except Exception as row_err:
                    stats["processed"] += 1
                    stats["failed"] += 1
                    if progress_callback:
                        progress_callback({
                            "index": item_num,
                            "total": len(valid_rows),
                            "row_idx": row_idx,
                            "card_code": card_code,
                            "zoho_id": zoho_id,
                            "syncflag": self.config.get("syncflag_value", "T"),
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

    # =========================================================================
    # MODO 2: ELIMINACIÓN MASIVA DE CLIENTES DUPLICADOS EN SAP B1
    # =========================================================================

    def run_delete_clients(
        self,
        progress_callback: Optional[Callable[[Dict[str, Any]], None]] = None,
        on_countdown_tick: Optional[Callable[[int], None]] = None,
        completion_callback: Optional[Callable[[Dict[str, Any]], None]] = None,
    ):
        """
        Ejecuta el ciclo de eliminación masiva en SAP B1.
        Para cada cliente:
          1. Modo Buscar en SAP (Ctrl + F)
          2. Digitar Código de Cliente + Enter
          3. Clic derecho en formulario / Menú Datos -> Eliminar
          4. Confirmar diálogo "¿Desea eliminar el socio de negocios?" con Enter / Alt+S
          5. Descartar avisos de bloqueo si tuviera transacciones
        """
        self._stop_event.clear()
        self._pause_event.set()
        self.is_running = True

        stats = {
            "total": 0,
            "processed": 0,
            "deleted": 0,
            "skipped_or_locked": 0,
            "failed": 0,
            "aborted": False,
            "error_msg": None,
        }

        try:
            hwnd, title = self._prepare_target_window(on_countdown_tick)

            df = self.dataframe
            sap_col = self.sap_col

            if df is None or sap_col not in df.columns:
                raise ValueError("El archivo de clientes a eliminar no ha sido cargado correctamente.")

            # Extraer lista de códigos únicos a eliminar
            client_codes = []
            for idx, row in df.iterrows():
                val = str(row[sap_col]).strip() if pd.notna(row[sap_col]) else ""
                if val and val.lower() != "nan" and val not in [c[1] for c in client_codes]:
                    client_codes.append((idx, val))

            stats["total"] = len(client_codes)

            for item_num, (row_idx, card_code) in enumerate(client_codes, start=1):
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
                        "total": len(client_codes),
                        "row_idx": row_idx,
                        "card_code": card_code,
                        "status": "BUSCANDO",
                        "message": f"Buscando cliente {card_code} en SAP...",
                    })

                try:
                    # PASO 1: Modo Buscar en SAP (Ctrl + F)
                    pyautogui.hotkey("ctrl", "f")
                    time.sleep(self.config.get("delay_after_find", 0.4))

                    # PASO 2: Digitar Código + Enter
                    self._safe_paste(card_code)
                    time.sleep(0.1)
                    pyautogui.press("enter")
                    time.sleep(self.config.get("delay_after_search", 0.9))

                    # PASO 3: Invocar acción "Eliminar"
                    delete_method = self.config.get("delete_method", "menu_alt_d")
                    right_click_coord = self.config.get("delete_click_coord")

                    if delete_method == "right_click" and right_click_coord:
                        # Clic derecho en la cabecera del socio de negocios
                        pyautogui.rightClick(right_click_coord[0], right_click_coord[1])
                        time.sleep(0.3)
                        # En el menú contextual de SAP B1 en español, 'e' activa "Eliminar"
                        pyautogui.press("e")
                    else:
                        # Método estándar por barra de menú SAP B1: Datos (Alt + D) -> Eliminar (e)
                        pyautogui.hotkey("alt", "d")
                        time.sleep(0.3)
                        pyautogui.press("e")

                    time.sleep(self.config.get("delay_after_delete_action", 0.5))

                    # PASO 4: Confirmar la ventana emergente de SAP ("¿Desea eliminar?")
                    pyautogui.press("enter")
                    time.sleep(self.config.get("delay_after_confirm_delete", 0.8))

                    # PASO 5: Descarte de avisos de error/bloqueo de SAP (ej. ligado a Cotización o Ventas)
                    # Si el socio tiene cotizaciones (OQUT) o ventas, SAP abre un modal de advertencia:
                    # "No se puede eliminar el socio de negocios; existen operaciones vinculadas"
                    # Pulsamos Enter y luego Escape para limpiar cualquier modal y dejar SAP listo para el siguiente
                    pyautogui.press("enter")
                    time.sleep(0.15)
                    pyautogui.press("escape")
                    time.sleep(0.2)

                    stats["processed"] += 1
                    stats["deleted"] += 1

                    if progress_callback:
                        progress_callback({
                            "index": item_num,
                            "total": len(client_codes),
                            "row_idx": row_idx,
                            "card_code": card_code,
                            "status": "PROCESADO",
                            "message": "Orden de eliminación enviada a SAP (si tiene coti/ventas, SAP lo protege)",
                        })

                except Exception as row_err:
                    stats["processed"] += 1
                    stats["failed"] += 1
                    if progress_callback:
                        progress_callback({
                            "index": item_num,
                            "total": len(client_codes),
                            "row_idx": row_idx,
                            "card_code": card_code,
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
