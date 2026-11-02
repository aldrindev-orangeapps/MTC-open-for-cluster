from __future__ import annotations

import ctypes
import os
import sys
from pathlib import Path


def bundled_template() -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return base / "assets" / "MTC Template.xlsx"


def settings_path() -> Path:
    return Path(os.environ.get("LOCALAPPDATA", Path.home())) / "MTCOpenOrder" / "settings.json"


def default_output_folder() -> Path:
    documents = Path.home() / "Documents"
    if sys.platform == "win32":
        buffer = ctypes.create_unicode_buffer(32768)
        if ctypes.windll.shell32.SHGetFolderPathW(None, 5, None, 0, buffer) == 0:
            documents = Path(buffer.value)
    return documents / "MTC Open Order" / "Outputs"
