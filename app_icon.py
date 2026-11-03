"""Set the same icon for the client and the server on Windows and macOS."""
import sys
from pathlib import Path
import tkinter as tk


def set_app_icon(root):
    assets = Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent))
    picture = assets / 'glazok.png'
    if picture.is_file():
        try:
            root._glazok_icon = tk.PhotoImage(master=root, file=str(picture))
            root.iconphoto(True, root._glazok_icon)
        except tk.TclError:
            pass
    icon = assets / '1.ico'
    if sys.platform == 'win32' and icon.is_file():
        try:
            root.iconbitmap(str(icon))
        except tk.TclError:
            pass
