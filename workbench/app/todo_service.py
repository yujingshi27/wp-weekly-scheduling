import calendar
import json
from datetime import date, datetime, timedelta

from db import get_conn

STATUS_TODO = "待完成"
STATUS_TODAY = "今日得做"
STATUS_BLOCKED = "卡着"
STATUS_WAITING = "等反馈"
STATUS_DONE = "已完成"

CATEGORY_REPEAT = "常规重复"
CATEGORIES = ["日常任务", "临时任务", CATEGORY_REPEAT, "DDL"]
DEFAULT_CATEGORY = "日常任务"
STATUSES = [STATUS_TODAY, STATUS_TODO, STATUS_BLOCKED, STATUS_WAITING, STATUS_DONE]
REPEAT_TYPES = ["none", "daily", "weekly", "monthly", "workday"]

QUADRANT_1 = "重要且紧急"
QUADRANT_2 = "重要不紧急"
QUADRANT_3 = "紧急不重要"
QUADRANT_4 = "不紧急不重要"
QUADRANTS = [QUADRANT_1, QUADRANT_2, QUADRANT_3, QUADRANT_4]
QUADRANT_RANK = {name: i for i, name in enumerate(QUADRANTS)}

ARCHIVE_KEEP_DAYS = 7


def _today():
    return date.today()


def _parse_date(value):
    if not value:
        return None
    if isinstance(value, date):
        return value
    try:
        return datetime.strptime(str(value)[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def _tags_for(conn, todo_id):
    rows = conn.execute(
        "SELECT t.name FROM tags t JOIN todo_tags tt ON tt.tag_id=t.id "
        "WHERE tt.todo_id=? ORDER BY t.name",
        (todo_id,),
    ).fetchall()
    return [r["name"] for r in rows]


def _row_to_dict(conn, row):
    todo = dict(row)
    todo["tags"] = _tags_for(conn, row["id"])
    due = _parse_date(row["due_date"])
    todo["overdue"] = bool(
        due and due < _today() and not row["archived_at"]
    )
    return todo


def ensure_default_tags():
    conn = get_conn()
    try:
        for name in ("能让 AI 做",):
            conn.execute("INSERT OR IGNORE INTO tags(name) VALUES (?)", (name,))
        conn.commit()
    finally:
        conn.close()


def list_tags():
    conn = get_conn()
    try:
        return [r["name"] for r in conn.execute("SELECT name FROM tags ORDER BY name")]
    finally:
        conn.close()


def add_tag(name):
    name = (name or "").strip()
    if not name:
        return None
    conn = get_conn()
    try:
        conn.execute("INSERT OR IGNORE INTO tags(name) VALUES (?)", (name,))
        conn.commit()
        row = conn.execute("SELECT id FROM tags WHERE name=?", (name,)).fetchone()
        return {"id": row["id"], "name": name}
    finally:
        conn.close()


def _set_tags(conn, todo_id, tags):
    conn.execute("DELETE FROM todo_tags WHERE todo_id=?", (todo_id,))
    for name in tags or []:
        name = (name or "").strip()
        if not name:
            continue
        conn.execute("INSERT OR IGNORE INTO tags(name) VALUES (?)", (name,))
        tag_id = conn.execute("SELECT id FROM tags WHERE name=?", (name,)).fetchone()["id"]
        conn.execute(
            "INSERT OR IGNORE INTO todo_tags(todo_id, tag_id) VALUES (?, ?)",
            (todo_id, tag_id),
        )


def apply_status_rules():
    """开机检查：把到期/过期的待办状态流转为“今日得做”。"""
    today = _today().isoformat()
    conn = get_conn()
    try:
        conn.execute(
            "UPDATE todos SET status=? "
            "WHERE archived_at IS NULL AND due_date IS NOT NULL "
            "AND date(due_date) <= date(?) "
            "AND status IN (?, ?, ?)",
            (STATUS_TODAY, today, STATUS_TODO, STATUS_BLOCKED, STATUS_WAITING),
        )
        conn.commit()
    finally:
        conn.close()


def purge_archived():
    """归档满 7 天彻底删除。"""
    cutoff = (_today() - timedelta(days=ARCHIVE_KEEP_DAYS)).isoformat()
    conn = get_conn()
    try:
        conn.execute(
            "DELETE FROM todos WHERE archived_at IS NOT NULL AND date(archived_at) <= date(?)",
            (cutoff,),
        )
        conn.commit()
    finally:
        conn.close()


def delete_expired_archive():
    purge_archived()


def create_todo(data):
    title = (data.get("title") or "").strip()
    if not title:
        raise ValueError("任务名称不能为空")
    category = data.get("category") or DEFAULT_CATEGORY
    if category not in CATEGORIES:
        category = DEFAULT_CATEGORY
    status = data.get("status") if data.get("status") in STATUSES else STATUS_TODO
    quadrant = data.get("quadrant") if data.get("quadrant") in QUADRANTS else None
    repeat_type = data.get("repeat_type") or "none"
    if category != CATEGORY_REPEAT:
        repeat_type = "none"
    if repeat_type not in REPEAT_TYPES:
        repeat_type = "none"
    repeat_param = data.get("repeat_param")
    if isinstance(repeat_param, (dict, list)):
        repeat_param = json.dumps(repeat_param, ensure_ascii=False)
    conn = get_conn()
    try:
        cur = conn.execute(
            "INSERT INTO todos(title, status, due_date, category, quadrant, "
            "repeat_type, repeat_param, note) VALUES (?,?,?,?,?,?,?,?)",
            (
                title,
                status,
                _norm_date(data.get("due_date")),
                category,
                quadrant,
                repeat_type,
                repeat_param,
                (data.get("note") or "").strip(),
            ),
        )
        todo_id = cur.lastrowid
        _set_tags(conn, todo_id, data.get("tags"))
        conn.commit()
    finally:
        conn.close()
    apply_status_rules()
    return get_todo(todo_id)


def _norm_date(value):
    d = _parse_date(value)
    return d.isoformat() if d else None


def get_todo(todo_id):
    conn = get_conn()
    try:
        row = conn.execute("SELECT * FROM todos WHERE id=?", (todo_id,)).fetchone()
        return _row_to_dict(conn, row) if row else None
    finally:
        conn.close()


def update_todo(todo_id, data):
    conn = get_conn()
    try:
        row = conn.execute("SELECT * FROM todos WHERE id=?", (todo_id,)).fetchone()
        if not row:
            return None
        fields = {}
        if "title" in data and (data["title"] or "").strip():
            fields["title"] = data["title"].strip()
        if "due_date" in data:
            fields["due_date"] = _norm_date(data["due_date"])
        if "note" in data:
            fields["note"] = (data.get("note") or "").strip()
        if "status" in data and data["status"] in STATUSES:
            fields["status"] = data["status"]
        if "category" in data:
            fields["category"] = data["category"] if data["category"] in CATEGORIES else DEFAULT_CATEGORY
        if "quadrant" in data:
            q = data["quadrant"]
            fields["quadrant"] = q if q in QUADRANTS else None
        if "repeat_type" in data:
            rt = data["repeat_type"] if data["repeat_type"] in REPEAT_TYPES else "none"
            fields["repeat_type"] = rt
        if "repeat_param" in data:
            rp = data["repeat_param"]
            fields["repeat_param"] = (
                json.dumps(rp, ensure_ascii=False) if isinstance(rp, (dict, list)) else rp
            )
        new_cat = fields.get("category", row["category"])
        if new_cat != CATEGORY_REPEAT:
            fields["repeat_type"] = "none"
        if fields:
            cols = ", ".join(f"{k}=?" for k in fields)
            conn.execute(f"UPDATE todos SET {cols} WHERE id=?", (*fields.values(), todo_id))
        if "tags" in data:
            _set_tags(conn, todo_id, data["tags"])
        conn.commit()
    finally:
        conn.close()
    apply_status_rules()
    return get_todo(todo_id)


def set_status(todo_id, status):
    if status not in STATUSES:
        raise ValueError("非法状态")
    return update_todo(todo_id, {"status": status})


def complete_todo(todo_id):
    conn = get_conn()
    try:
        row = conn.execute("SELECT * FROM todos WHERE id=?", (todo_id,)).fetchone()
        if not row:
            return None
        now = datetime.now().isoformat(timespec="seconds")
        conn.execute(
            "UPDATE todos SET status=?, completed_at=?, archived_at=? WHERE id=?",
            (STATUS_DONE, now, now, todo_id),
        )
        conn.commit()
        spawn = _spawn_next(conn, row)
        conn.commit()
        return {"completed": todo_id, "next": spawn}
    finally:
        conn.close()


def _advance(due, repeat_type, param):
    param = param or {}
    if repeat_type == "daily":
        return due + timedelta(days=1)
    if repeat_type == "workday":
        d = due + timedelta(days=1)
        while d.weekday() >= 5:
            d += timedelta(days=1)
        return d
    if repeat_type == "weekly":
        days = param.get("weekdays") or []
        days = sorted(int(x) for x in days)
        if not days:
            return due + timedelta(days=7)
        d = due + timedelta(days=1)
        for _ in range(14):
            if d.weekday() in days:
                return d
            d += timedelta(days=1)
        return due + timedelta(days=7)
    if repeat_type == "monthly":
        day = int(param.get("day") or due.day)
        year = due.year + (1 if due.month == 12 else 0)
        month = 1 if due.month == 12 else due.month + 1
        last = calendar.monthrange(year, month)[1]
        return date(year, month, min(day, last))
    return None


def _spawn_next(conn, row):
    if row["category"] != CATEGORY_REPEAT or row["repeat_type"] in (None, "none"):
        return None
    due = _parse_date(row["due_date"])
    if not due:
        return None
    try:
        param = json.loads(row["repeat_param"]) if row["repeat_param"] else {}
    except (json.JSONDecodeError, TypeError):
        param = {}
    next_due = _advance(due, row["repeat_type"], param)
    if not next_due:
        return None
    cur = conn.execute(
        "INSERT INTO todos(title, status, due_date, category, quadrant, "
        "repeat_type, repeat_param, note) VALUES (?,?,?,?,?,?,?,?)",
        (
            row["title"],
            STATUS_TODO,
            next_due.isoformat(),
            row["category"],
            row["quadrant"],
            row["repeat_type"],
            row["repeat_param"],
            row["note"],
        ),
    )
    new_id = cur.lastrowid
    for name in _tags_for(conn, row["id"]):
        conn.execute("INSERT OR IGNORE INTO tags(name) VALUES (?)", (name,))
        tag_id = conn.execute("SELECT id FROM tags WHERE name=?", (name,)).fetchone()["id"]
        conn.execute(
            "INSERT OR IGNORE INTO todo_tags(todo_id, tag_id) VALUES (?, ?)", (new_id, tag_id)
        )
    return new_id


def delete_todo(todo_id):
    conn = get_conn()
    try:
        conn.execute("DELETE FROM todos WHERE id=?", (todo_id,))
        conn.execute("DELETE FROM todo_tags WHERE todo_id=?", (todo_id,))
        conn.commit()
        return True
    finally:
        conn.close()


def list_active_by_status():
    conn = get_conn()
    try:
        rank = " ".join(
            f"WHEN '{q}' THEN {i}" for q, i in QUADRANT_RANK.items()
        )
        rows = conn.execute(
            "SELECT * FROM todos WHERE archived_at IS NULL "
            "ORDER BY (due_date IS NULL), date(due_date) ASC, "
            f"CASE quadrant {rank} ELSE {len(QUADRANTS)} END ASC, id ASC"
        ).fetchall()
        groups = {s: [] for s in [STATUS_TODAY, STATUS_TODO, STATUS_BLOCKED, STATUS_WAITING]}
        for row in rows:
            todo = _row_to_dict(conn, row)
            groups.setdefault(row["status"], []).append(todo)
        return groups
    finally:
        conn.close()


def list_archived(keyword=None):
    conn = get_conn()
    try:
        sql = "SELECT * FROM todos WHERE archived_at IS NOT NULL"
        params = []
        if keyword:
            sql += " AND title LIKE ?"
            params.append(f"%{keyword}%")
        sql += " ORDER BY datetime(archived_at) DESC, id DESC"
        rows = conn.execute(sql, params).fetchall()
        return [_row_to_dict(conn, r) for r in rows]
    finally:
        conn.close()
