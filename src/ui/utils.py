import os
import sys
from screeninfo import get_monitors
import flet as ft

def get_app_size():
    monitor = None
    monitors = get_monitors()
    
    for m in monitors:
        if m.is_primary:
            monitor = m
            break
    
    if monitor:
        app_width = monitor.width * 0.40
        app_height = monitor.height * 0.60
        
        return app_width, app_height
    else:
        print("Não foi possível encontrar o monitor primário.")

def get_icon_path():
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, 'assets', 'icon_bosch.ico')
    else:
        return os.path.join(os.path.dirname(__file__), 'assets', 'icon_bosch.ico')

def config(page):
    assets_path = get_icon_path()
    page.bgcolor ="#FFFFFF"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.title = "PDF-Reader"
    app_width, app_height = get_app_size()
    page.window.width = app_width
    page.window.height = app_height
    page.window.maximizable = False
    page.window.resizable = False
    page.padding=0
    page.spacing=0
    page.scroll=None
    page.window_icon = assets_path
    page.window.center()
    page.update()
