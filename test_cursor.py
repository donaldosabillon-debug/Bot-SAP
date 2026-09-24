import ctypes
import win32api
import win32con
import win32service

ctypes.windll.shcore.SetProcessDpiAwareness(2)

print("Virtual Screen:")
print("SM_XVIRTUALSCREEN:", win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN))
print("SM_YVIRTUALSCREEN:", win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN))
print("SM_CXVIRTUALSCREEN:", win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN))
print("SM_CYVIRTUALSCREEN:", win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN))

desk = win32service.OpenDesktop("default", 0, False, win32con.GENERIC_ALL)
desk.SetThreadDesktop()

# Try setting to (500, 500)
r1 = ctypes.windll.user32.SetCursorPos(500, 500)
print("SetCursorPos(500, 500):", r1, "GetLastError:", ctypes.GetLastError())

# Try setting to (2500, 300)
r2 = ctypes.windll.user32.SetCursorPos(2500, 300)
print("SetCursorPos(2500, 300):", r2, "GetLastError:", ctypes.GetLastError())
