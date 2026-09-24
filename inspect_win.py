import ctypes
import win32gui
import win32con

ctypes.windll.shcore.SetProcessDpiAwareness(2)

def callback(hwnd, extra):
    if win32gui.IsWindowVisible(hwnd):
        title = win32gui.GetWindowText(hwnd)
        if title:
            rect = win32gui.GetWindowRect(hwnd)
            print(f"[{hwnd}] Rect: {rect} | Title: {title}")
    return True

win32gui.EnumWindows(callback, None)
