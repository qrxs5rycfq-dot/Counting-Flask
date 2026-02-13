import os
import sys

def get_base_dir():
    if getattr(sys, 'frozen', False):
        # Saat dibundle oleh PyInstaller, gunakan folder _MEIPASS
        return sys._MEIPASS
    return os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))