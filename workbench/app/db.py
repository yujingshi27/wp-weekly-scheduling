import sqlite3
import threading
from pathlib import Path

from config import DATA_DIR, ensure_dirs

DB_PATH = DATA_DIR / "todo.db"

JOB_LOCK = threading.Lock()

SCHEMA = """
CREATE TABLE IF NOT EXISTS todos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    title TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT '待完成',
    due_date TEXT,
    category TEXT NOT NULL DEFAULT '日常任务',
    quadrant TEXT,
    repeat_type TEXT NOT NULL DEFAULT 'none',
    repeat_param TEXT,
    note TEXT,
    source_code TEXT,
    source_node TEXT,
    completed_at TEXT,
    archived_at TEXT
);

CREATE TABLE IF NOT EXISTS tags (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS todo_tags (
    todo_id INTEGER NOT NULL,
    tag_id INTEGER NOT NULL,
    PRIMARY KEY (todo_id, tag_id)
);

CREATE TABLE IF NOT EXISTS ts_rows (
    project_code TEXT,
    project_name TEXT,
    staff TEXT,
    manager TEXT,
    hours REAL,
    period TEXT,
    description TEXT
);

CREATE TABLE IF NOT EXISTS wp_budget (
    project_code TEXT PRIMARY KEY,
    project_name TEXT,
    outlook_hours REAL,
    updated_at TEXT
);

CREATE TABLE IF NOT EXISTS meta (
    key TEXT PRIMARY KEY,
    value TEXT
);

CREATE INDEX IF NOT EXISTS idx_todos_status ON todos(status);
CREATE INDEX IF NOT EXISTS idx_todos_source ON todos(source_code, source_node);
CREATE INDEX IF NOT EXISTS idx_ts_code ON ts_rows(project_code);
CREATE INDEX IF NOT EXISTS idx_ts_manager ON ts_rows(manager);
"""

DEFAULT_TAGS = ["能让 AI 做"]


def get_conn():
    ensure_dirs()
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA busy_timeout = 30000")
    return conn


def init_db():
    conn = get_conn()
    try:
        conn.execute("PRAGMA journal_mode = WAL")
        conn.executescript(SCHEMA)
        _migrate(conn)
        for name in DEFAULT_TAGS:
            conn.execute("INSERT OR IGNORE INTO tags(name) VALUES (?)", (name,))
        conn.commit()
    finally:
        conn.close()


def _migrate(conn):
    cols = {row["name"] for row in conn.execute("PRAGMA table_info(todos)")}
    if "quadrant" not in cols:
        conn.execute("ALTER TABLE todos ADD COLUMN quadrant TEXT")
    conn.execute("UPDATE todos SET category='日常任务' WHERE category='其他'")
