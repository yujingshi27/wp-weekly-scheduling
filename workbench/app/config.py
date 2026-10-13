import json
import sys
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = APP_DIR.parent

DATA_DIR = BASE_DIR / "data"
BACKUP_DIR = BASE_DIR / "backup"
EXPORT_DIR = BASE_DIR / "exports"
CONFIG_PATH = BASE_DIR / "config.json"


def resource_path(rel):
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS) / rel
    return APP_DIR / rel

DEFAULT_CONFIG = {
    "datasources": {
        "timesheet": "",
        "wp_service": "",
        "schedule_output": "",
    },
    "timesheet_mapping": {
        "project_code": "Engagement Code",
        "staff": "Employee Name",
        "manager": "COE Manager",
        "hours": "Charge Hours",
        "period": "Transaction Cycle Date",
        "description": "Charge Description",
        "project_name": "Engagement Name",
    },
    "timesheet_sheet": "Timesheet",
    "last_backup_week": None,
    "port": 5000,
}


def ensure_dirs():
    for d in (DATA_DIR, BACKUP_DIR, EXPORT_DIR):
        d.mkdir(parents=True, exist_ok=True)


def load_config():
    ensure_dirs()
    if not CONFIG_PATH.exists():
        save_config(DEFAULT_CONFIG)
        return json.loads(json.dumps(DEFAULT_CONFIG))
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    except (json.JSONDecodeError, OSError):
        cfg = {}
    merged = json.loads(json.dumps(DEFAULT_CONFIG))
    for k, v in cfg.items():
        if isinstance(v, dict) and isinstance(merged.get(k), dict):
            merged[k].update(v)
        else:
            merged[k] = v
    return merged


def save_config(cfg):
    ensure_dirs()
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
