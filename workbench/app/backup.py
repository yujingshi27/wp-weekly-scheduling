import sqlite3
from datetime import date

from config import BACKUP_DIR, load_config, save_config
from db import DB_PATH

KEEP = 7
PREFIX = "todo_"


def _current_week_key():
    y, w, _ = date.today().isocalendar()
    return f"{y}-W{w:02d}"


def _prune():
    files = sorted(BACKUP_DIR.glob(PREFIX + "*.db"), key=lambda p: p.name, reverse=True)
    for old in files[KEEP:]:
        try:
            old.unlink()
        except OSError:
            pass


def run_weekly_backup(force=False):
    cfg = load_config()
    week = _current_week_key()
    if not force and cfg.get("last_backup_week") == week:
        return {"backed_up": False, "reason": "本周已备份"}
    if not DB_PATH.exists():
        return {"backed_up": False, "reason": "数据库尚未创建"}
    dest = BACKUP_DIR / f"{PREFIX}{date.today().strftime('%Y%m%d')}.db"
    src = sqlite3.connect(DB_PATH, timeout=30)
    try:
        dst = sqlite3.connect(dest)
        try:
            with dst:
                src.backup(dst)
        finally:
            dst.close()
    finally:
        src.close()
    _prune()
    cfg["last_backup_week"] = week
    save_config(cfg)
    return {"backed_up": True, "path": str(dest)}


def list_backups():
    files = sorted(BACKUP_DIR.glob(PREFIX + "*.db"), key=lambda p: p.name, reverse=True)
    return [{"name": p.name, "size": p.stat().st_size} for p in files]
