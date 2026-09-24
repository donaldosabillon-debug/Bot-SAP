import ctypes
import pyautogui
import win32service
import win32con

ctypes.windll.shcore.SetProcessDpiAwareness(2)
pyautogui.FAILSAFE = False

desk = win32service.OpenDesktop("default", 0, False, win32con.GENERIC_ALL)
desk.SetThreadDesktop()

print("pyautogui size:", pyautogui.size())
try:
    pyautogui.moveTo(2500, 300)
    print("pyautogui position after moveTo(2500, 300):", pyautogui.position())
except Exception as e:
    print("pyautogui error:", e)
