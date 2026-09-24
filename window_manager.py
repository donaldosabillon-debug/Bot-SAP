"""
window_manager.py
Módulo robusto de detección, listado y activación de ventanas en Windows.
Compatible con sesiones interactivas y entornos de terminal con soporte de SetThreadDesktop.
Especializado en detectar sesiones de Escritorio Remoto (mstsc) y SAP Business One.
"""

import ctypes
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2) # Per-monitor DPI aware
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

import time
import win32gui
import win32con
import win32api
import win32service
from typing import List, Tuple, Optional

KEYWORDS_SAP_RDP = [
    "182.160",
    "conexión a escritorio remoto",
    "conexion a escritorio remoto",
    "escritorio remoto",
    "remote desktop",
    "sap business one",
    "sap b1",
    "sap hana",
    "sap logon",
]

_DESKTOP_HANDLE = None


def ensure_desktop_attached():
    """
    Asegura que el hilo actual esté vinculado al escritorio 'default' interactivo de Windows.
    Esto permite enumerar e interactuar con ventanas reales del usuario incluso si el script
    inicia desde una terminal o entorno de ejecución secundario.
    """
    global _DESKTOP_HANDLE
    try:
        if _DESKTOP_HANDLE is None:
            _DESKTOP_HANDLE = win32service.OpenDesktop("default", 0, False, win32con.GENERIC_ALL)
            if _DESKTOP_HANDLE:
                _DESKTOP_HANDLE.SetThreadDesktop()
    except Exception as e:
        pass


def list_visible_windows() -> List[Tuple[int, str]]:
    """
    Retorna lista de (hwnd, titulo) de todas las ventanas principales visibles en el escritorio interactivo.
    """
    ensure_desktop_attached()
    results = []
    seen_hwnds = set()

    def process_hwnd(hwnd):
        if hwnd in seen_hwnds:
            return
        seen_hwnds.add(hwnd)
        try:
            if win32gui.IsWindow(hwnd):
                title = win32gui.GetWindowText(hwnd).strip()
                if not title or title in ("Program Manager", "Default IME", "MSCTFIME UI"):
                    return
                # Ignorar WhatsApp explícitamente para evitar falsos positivos con "sap"
                if "whatsapp" in title.lower():
                    return
                is_iconic = win32gui.IsIconic(hwnd)
                is_visible = win32gui.IsWindowVisible(hwnd)
                if is_visible or is_iconic:
                    rect = win32gui.GetWindowRect(hwnd)
                    w = rect[2] - rect[0]
                    h_len = rect[3] - rect[1]
                    if (w > 50 and h_len > 50) or is_iconic:
                        results.append((hwnd, title))
        except Exception:
            pass

    # 1. Intentar a través del desktop interactivo
    try:
        if _DESKTOP_HANDLE:
            for h in _DESKTOP_HANDLE.EnumDesktopWindows():
                process_hwnd(int(h))
    except Exception:
        pass

    # 2. Complementar con EnumWindows estándar
    def enum_cb(hwnd, _):
        process_hwnd(hwnd)
        return True

    try:
        win32gui.EnumWindows(enum_cb, None)
    except Exception:
        pass

    # Ordenar alfabéticamente por título
    results.sort(key=lambda x: x[1].lower())
    return results


def find_target_window(custom_title: Optional[str] = None) -> Optional[Tuple[int, str]]:
    """
    Busca la ventana de SAP o Escritorio Remoto activa.
    1. Si se provee custom_title, busca coincidencia parcial.
    2. Si no, busca por palabras clave conocidas (Escritorio Remoto, SAP, etc.).
    Retorna (hwnd, titulo) o None.
    """
    visible = list_visible_windows()

    # 1. Búsqueda por título personalizado estricto
    if custom_title and custom_title.strip():
        target_lower = custom_title.lower().strip()
        for hwnd, title in visible:
            if target_lower in title.lower():
                return hwnd, title
        return None

    # 2. Búsqueda por palabras clave con prioridad
    # Prioridad A: Escritorio Remoto con conexión activa (182.160... o Escritorio Remoto)
    for hwnd, title in visible:
        t_low = title.lower()
        if "whatsapp" in t_low:
            continue
        if "182.160" in t_low or "conexión a escritorio remoto" in t_low or "conexion a escritorio remoto" in t_low or "escritorio remoto" in t_low or "remote desktop" in t_low:
            return hwnd, title

    # Prioridad B: SAP Business One directo
    for hwnd, title in visible:
        t_low = title.lower()
        if "whatsapp" in t_low:
            continue
        if "sap business one" in t_low or "sap b1" in t_low or "sap hana" in t_low:
            return hwnd, title

    # Prioridad C: Cualquier otra palabra clave de la lista
    for hwnd, title in visible:
        t_low = title.lower()
        if "whatsapp" in t_low:
            continue
        for kw in KEYWORDS_SAP_RDP:
            if kw in t_low:
                return hwnd, title

    return None


def is_window_valid(hwnd: int) -> bool:
    """Verifica si un HWND sigue siendo una ventana válida existente."""
    try:
        return bool(win32gui.IsWindow(hwnd))
    except Exception:
        return False


def bring_window_to_front(hwnd: int) -> bool:
    """
    Restaura la ventana y le otorga el foco en primer plano de forma garantizada en Windows 10/11.
    """
    ensure_desktop_attached()
    try:
        if not win32gui.IsWindow(hwnd):
            return False

        # Si está minimizada, restaurar
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        else:
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)

        # Liberar la restricción de Windows para cambiar de ventana activa
        win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)
        win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)

        win32gui.BringWindowToTop(hwnd)
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.4)
        return True
    except Exception as e:
        print(f"[window_manager] Error activando ventana {hwnd}: {e}")
        return False


def get_window_bounds(hwnd: int) -> Optional[Tuple[int, int, int, int]]:
    """Retorna (x, y, width, height) de la ventana."""
    try:
        if win32gui.IsWindow(hwnd):
            rect = win32gui.GetWindowRect(hwnd)
            return rect[0], rect[1], rect[2] - rect[0], rect[3] - rect[1]
    except Exception:
        pass
    return None
