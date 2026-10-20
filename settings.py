"""%APPDATA%\\MD문서변환기\\settings.json 에 마지막 설정을 저장·복원한다."""
from __future__ import annotations

import json
import os

APP_DIR_NAME = "MD문서변환기"

DEFAULTS = {
    "out_dir": "",
    "same_folder": True,
    "format": "docx",          # docx | pdf
    "auto_number_bold": True,
    "auto_box": True,
    "geometry": None,
}


def settings_path() -> str:
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(base, APP_DIR_NAME, "settings.json")


def load() -> dict:
    data = dict(DEFAULTS)
    try:
        with open(settings_path(), encoding="utf-8") as f:
            saved = json.load(f)
        if isinstance(saved, dict):
            data.update({k: v for k, v in saved.items() if k in DEFAULTS})
    except (OSError, ValueError):
        pass
    return data


def save(data: dict):
    path = settings_path()
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump({k: data.get(k, v) for k, v in DEFAULTS.items()}, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
    except OSError:
        pass
