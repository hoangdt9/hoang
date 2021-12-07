import os
from winreg import *

def get_current_zoom():
    try:
        REG_PATH = 'Software\Microsoft\Internet Explorer\Zoom'
        hreg = ConnectRegistry(None, HKEY_CURRENT_USER)
        hkey = OpenKey(hreg, REG_PATH, 0,KEY_READ)
        current_zoom_factor = QueryValueEx(hkey, 'ZoomFactor')
        CloseKey(hkey)
        return current_zoom_factor
    except WindowsError:
        return None

def get_dpi():
    try:
        REG_PATH = 'Control Panel\Desktop\WindowMetrics'
        hreg = ConnectRegistry(None, HKEY_CURRENT_USER)
        hkey = OpenKey(hreg, REG_PATH, 0,KEY_READ)
        dpi = QueryValueEx(hkey, 'AppliedDPI')
        CloseKey(hkey)
        return dpi
    except WindowsError:
        return None

def set_zoom_100():
    try:
        keyval = r"Software\Microsoft\Internet Explorer\Zoom"
        if not os.path.exists("keyval"):
            key = CreateKey(HKEY_CURRENT_USER, keyval)

        dpi = get_dpi()
        zoom_factor_100_percent = 100000
        if dpi[0] == 120:
            zoom_factor_100_percent = 80000
        elif dpi[0] == 144:
            zoom_factor_100_percent = 66667

        current_zoom = get_current_zoom()[0]
        if current_zoom != zoom_factor_100_percent:
            hkey = OpenKey(HKEY_CURRENT_USER, keyval, 0, KEY_WRITE)
            SetValueEx(hkey, 'ZoomFactor', 0, REG_DWORD, zoom_factor_100_percent)
            print('set zoom to 100')
            CloseKey(hkey)
        return True
    except WindowsError:
        return False