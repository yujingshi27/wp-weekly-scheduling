from collections import OrderedDict

from db import get_conn


def norm_code(value):
    if value is None:
        return ""
    s = str(value).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def clean_list(values):
    out = []
    for v in values or []:
        s = (v or "").strip() if isinstance(v, str) else v
        if s not in (None, ""):
            out.append(str(s).strip())
    return out


def filter_records(mode, codes, managers):
    codes = sorted({norm_code(c) for c in clean_list(codes)})
    managers = clean_list(managers)
    where = []
    params = []
    if mode == "project" and codes:
        where.append("project_code IN (" + ",".join("?" * len(codes)) + ")")
        params.extend(codes)
    if managers:
        sub = " OR ".join("LOWER(manager) LIKE ?" for _ in managers)
        where.append("(" + sub + ")")
        params.extend("%" + m.lower() + "%" for m in managers)
    sql = (
        "SELECT project_code, project_name, staff, manager, hours, period, description "
        "FROM ts_rows"
    )
    if where:
        sql += " WHERE " + " AND ".join(where)
    conn = get_conn()
    try:
        rows = conn.execute(sql, params).fetchall()
    finally:
        conn.close()
    return [dict(r) for r in rows]


def aggregate(rows):
    agg = OrderedDict()
    for r in rows:
        key = (
            r.get("project_name") or "",
            norm_code(r.get("project_code")),
            r.get("staff") or "",
            r.get("manager") or "",
            r.get("period") or "",
            r.get("description") or "",
        )
        agg[key] = agg.get(key, 0.0) + float(r.get("hours") or 0)
    result = []
    for (pn, code, staff, mgr, period, desc), hours in agg.items():
        result.append({
            "project_name": pn,
            "project_code": code,
            "staff": staff,
            "manager": mgr,
            "hours": round(hours, 4),
            "period": period,
            "description": desc,
        })
    result.sort(key=lambda x: (x["project_code"], x["staff"], x["period"], x["description"]))
    return result


def get_budgets():
    conn = get_conn()
    try:
        rows = conn.execute(
            "SELECT project_code, project_name, outlook_hours FROM wp_budget"
        ).fetchall()
    finally:
        conn.close()
    budgets = {}
    for r in rows:
        budgets[norm_code(r["project_code"])] = {
            "project_name": r["project_name"],
            "outlook_hours": r["outlook_hours"],
        }
    return budgets


def budget_count():
    conn = get_conn()
    try:
        row = conn.execute("SELECT COUNT(*) AS c FROM wp_budget").fetchone()
        return row["c"]
    finally:
        conn.close()


def ts_row_count():
    conn = get_conn()
    try:
        row = conn.execute("SELECT COUNT(*) AS c FROM ts_rows").fetchone()
        return row["c"]
    finally:
        conn.close()


def project_totals(rows):
    budgets = get_budgets()
    totals = OrderedDict()
    for r in rows:
        code = norm_code(r.get("project_code"))
        if code not in totals:
            totals[code] = {
                "project_code": code,
                "project_name": r.get("project_name") or "",
                "hours": 0.0,
            }
        totals[code]["hours"] += float(r.get("hours") or 0)
    result = []
    for code in sorted(totals.keys()):
        t = totals[code]
        b = budgets.get(code)
        budget = b["outlook_hours"] if b else None
        t["budget"] = budget
        t["over"] = bool(budget is not None and t["hours"] > float(budget))
        t["hours"] = round(t["hours"], 4)
        result.append(t)
    return result


def build_result(mode, codes, managers, limit=None):
    rows = filter_records(mode, codes, managers)
    agg = aggregate(rows)
    row_count = len(agg)
    truncated = False
    if limit and row_count > limit:
        agg = agg[:limit]
        truncated = True
    return {
        "rows": agg,
        "totals": project_totals(rows),
        "raw_count": len(rows),
        "row_count": row_count,
        "truncated": truncated,
        "budget_count": budget_count(),
    }
