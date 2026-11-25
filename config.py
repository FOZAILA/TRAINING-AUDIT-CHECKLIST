import json
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SETTINGS_PATH = os.path.join(BASE_DIR, "config", "settings.json")
FORMATS_DIR = os.path.join(BASE_DIR, "audit_formats")
DATA_DIR = os.path.join(BASE_DIR, "data")


def load_settings():
    """Read settings.json fresh every time, so edits apply without restarting."""
    with open(SETTINGS_PATH, encoding="utf-8") as fh:
        return json.load(fh)
