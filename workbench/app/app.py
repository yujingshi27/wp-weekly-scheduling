import os
import socket
import subprocess
import sys
import threading
import traceback
import urllib.request
import webbrowser
from datetime import datetime
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.exceptions import HTTPException

sys.path.insert(0, str(Path(__file__).resolve().parent))

import datasource as ds  # noqa: E402
import backup as backup_mod  # noqa: E402
import ddl_service as ddl  # noqa: E402
import export as export_mod  # noqa: E402
import todo_service as todo  # noqa: E402
import timesheet_service as ts  # noqa: E402
from config import load_config, ensure_dirs, resource_path, BASE_DIR  # noqa: E402
from db import init_db  # noqa: E402

APP_ID = "workbench"

app = Flask(
    __name__,
    template_folder=str(resource_path("templates")),
    static_folder=str(resource_path("static")),
)
app.config["JSON_AS_ASCII"] = False


def _log_error(text):
    try:
        log_dir = BASE_DIR / "logs"
        log_dir.mkdir(parents=True, exist_ok=True)
        with open(log_dir / "error.log", "a", encoding="utf-8") as f:
            f.write(f"\n[{datetime.now().isoformat(timespec='seconds')}]\n{text}\n")
    except OSError:
        pass


@app.errorhandler(Exception)
def handle_uncaught(exc):
    if isinstance(exc, HTTPException):
        return exc
    _log_error(traceback.format_exc())
    return jsonify({"ok": False, "error": f"{type(exc).__name__}: {exc}"}), 500


@app.context_processor
def inject_globals():
    return {"app_name": "WP 工作台"}


@app.route("/api/health")
def api_health():
    return jsonify({"ok": True, "app": APP_ID})


@app.route("/")
def index():
    return render_template("index.html", active="home")


@app.route("/timesheet")
def timesheet():
    return render_template("timesheet.html", active="timesheet")


@app.route("/settings")
def settings():
    return render_template("settings.html", active="settings")


@app.route("/api/todos", methods=["GET"])
def api_list_todos():
    groups = todo.list_active_by_status()
    counts = {k: len(v) for k, v in groups.items()}
    return jsonify({"ok": True, "groups": groups, "counts": counts})


@app.route("/api/todos/archive", methods=["GET"])
def api_list_archive():
    q = request.args.get("q", "").strip() or None
    return jsonify({"ok": True, "items": todo.list_archived(q)})


@app.route("/api/todos", methods=["POST"])
def api_create_todo():
    try:
        item = todo.create_todo(request.get_json(force=True) or {})
        return jsonify({"ok": True, "item": item})
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.route("/api/todos/<int:todo_id>", methods=["GET"])
def api_get_todo(todo_id):
    item = todo.get_todo(todo_id)
    if not item:
        return jsonify({"ok": False, "error": "待办不存在"}), 404
    return jsonify({"ok": True, "item": item})


@app.route("/api/todos/<int:todo_id>", methods=["PATCH"])
def api_update_todo(todo_id):
    item = todo.update_todo(todo_id, request.get_json(force=True) or {})
    if not item:
        return jsonify({"ok": False, "error": "待办不存在"}), 404
    return jsonify({"ok": True, "item": item})


@app.route("/api/todos/<int:todo_id>", methods=["DELETE"])
def api_delete_todo(todo_id):
    todo.delete_todo(todo_id)
    return jsonify({"ok": True})


@app.route("/api/todos/<int:todo_id>/complete", methods=["POST"])
def api_complete_todo(todo_id):
    result = todo.complete_todo(todo_id)
    if not result:
        return jsonify({"ok": False, "error": "待办不存在"}), 404
    return jsonify({"ok": True, "result": result})


@app.route("/api/tags", methods=["GET"])
def api_list_tags():
    return jsonify({"ok": True, "tags": todo.list_tags()})


@app.route("/api/tags", methods=["POST"])
def api_add_tag():
    tag = todo.add_tag((request.get_json(force=True) or {}).get("name"))
    if not tag:
        return jsonify({"ok": False, "error": "标签名不能为空"}), 400
    return jsonify({"ok": True, "tag": tag})


@app.route("/api/timesheet/filter", methods=["POST"])
def api_ts_filter():
    data = request.get_json(force=True) or {}
    mode = data.get("mode") or "project"
    if mode not in ("project", "person"):
        mode = "project"
    result = ts.build_result(mode, data.get("codes"), data.get("managers"), limit=2000)
    return jsonify({"ok": True, **result})


@app.route("/api/timesheet/status", methods=["GET"])
def api_ts_status():
    cfg = load_config()
    return jsonify({
        "ok": True,
        "ts_rows": ts.ts_row_count(),
        "budgets": ts.budget_count(),
        "datasources": cfg.get("datasources", {}),
        "datasource_exists": ds.verify_datasources(),
        "ts_last_import": ds.get_meta("ts_last_import"),
        "budget_last_import": ds.get_meta("budget_last_import"),
    })


@app.route("/api/settings", methods=["GET"])
def api_get_settings():
    cfg = load_config()
    return jsonify({
        "ok": True,
        "datasources": cfg.get("datasources", {}),
        "timesheet_mapping": cfg.get("timesheet_mapping", {}),
        "timesheet_sheet": cfg.get("timesheet_sheet", ""),
        "mapping_fields": ds.TIMESHEET_KEYS,
    })


@app.route("/api/settings", methods=["POST"])
def api_save_settings():
    data = request.get_json(force=True) or {}
    cfg = load_config()
    if "datasources" in data:
        for k, v in (data["datasources"] or {}).items():
            if k in cfg["datasources"]:
                cfg["datasources"][k] = (v or "").strip()
    if "timesheet_mapping" in data:
        for k, v in (data["timesheet_mapping"] or {}).items():
            if k in cfg["timesheet_mapping"]:
                cfg["timesheet_mapping"][k] = (v or "").strip()
    if "timesheet_sheet" in data:
        cfg["timesheet_sheet"] = (data["timesheet_sheet"] or "").strip()
    save_config(cfg)
    return jsonify({"ok": True})


@app.route("/api/settings/sheets", methods=["POST"])
def api_settings_sheets():
    data = request.get_json(force=True) or {}
    source = data.get("source")
    cfg = load_config()
    path = (data.get("path") or cfg["datasources"].get(source) or "").strip()
    try:
        return jsonify({"ok": True, "sheets": ds.list_sheets(path)})
    except (FileNotFoundError, OSError) as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.route("/api/settings/headers", methods=["POST"])
def api_settings_headers():
    data = request.get_json(force=True) or {}
    source = data.get("source")
    cfg = load_config()
    path = (data.get("path") or cfg["datasources"].get(source) or "").strip()
    sheet = data.get("sheet") or None
    try:
        return jsonify({"ok": True, "headers": ds.read_headers(path, sheet)})
    except (FileNotFoundError, OSError) as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.route("/api/timesheet/refresh", methods=["POST"])
def api_ts_refresh():
    try:
        return jsonify({"ok": True, **ds.import_timesheet()})
    except (FileNotFoundError, ValueError, OSError) as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.route("/api/timesheet/refresh_budget", methods=["POST"])
def api_ts_refresh_budget():
    try:
        return jsonify({"ok": True, **ds.import_budget()})
    except (FileNotFoundError, ValueError, OSError) as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.route("/api/pick_file", methods=["POST"])
def api_pick_file():
    data = request.get_json(silent=True) or {}
    title = (data.get("title") or "选择数据源文件").replace("'", "''")
    ps = (
        "$OutputEncoding = [System.Text.Encoding]::UTF8;"
        "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8;"
        "Add-Type -AssemblyName System.Windows.Forms;"
        "$d = New-Object System.Windows.Forms.OpenFileDialog;"
        f"$d.Title = '{title}';"
        "$d.Filter = 'Excel 文件 (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|所有文件 (*.*)|*.*';"
        "$d.CheckFileExists = $true;"
        "if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) "
        "{ [Console]::Out.Write($d.FileName) }"
    )
    try:
        out = subprocess.run(
            ["powershell", "-NoProfile", "-STA", "-Command", ps],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=600,
        )
        return jsonify({"ok": True, "path": (out.stdout or "").strip()})
    except (subprocess.SubprocessError, OSError) as exc:
        return jsonify({"ok": False, "error": "无法打开文件选择框：" + str(exc)}), 400


@app.route("/api/timesheet/export", methods=["POST"])
def api_ts_export():
    data = request.get_json(force=True) or {}
    mode = data.get("mode") or "project"
    if mode not in ("project", "person"):
        mode = "project"
    try:
        info = export_mod.export_timesheet(mode, data.get("codes"), data.get("managers"))
    except (ValueError, OSError) as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400
    return send_file(info["path"], as_attachment=True, download_name=info["filename"])


@app.route("/api/ddl/refresh", methods=["POST"])
def api_ddl_refresh():
    try:
        return jsonify({"ok": True, **ddl.refresh_ddl()})
    except (FileNotFoundError, ValueError, OSError) as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


def _is_our_app(port):
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/api/health", timeout=0.8) as resp:
            body = resp.read().decode("utf-8", "ignore")
        return '"app": "workbench"' in body or '"app":"workbench"' in body
    except Exception:
        return False


def _resolve_port(start, span=20):
    """返回 (port, already_running)。已在运行则 port 为现有端口。"""
    for port in range(start, start + span):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            in_use = s.connect_ex(("127.0.0.1", port)) == 0
        if not in_use:
            return port, False
        if _is_our_app(port):
            return port, True
    return start, False


def _open_browser(url):
    if os.environ.get("WORKBENCH_NO_BROWSER") == "1":
        return
    threading.Timer(1.0, lambda: webbrowser.open(url)).start()


def main():
    ensure_dirs()
    init_db()
    todo.ensure_default_tags()
    todo.apply_status_rules()
    todo.purge_archived()
    try:
        result = backup_mod.run_weekly_backup()
        if result.get("backed_up"):
            print("已自动备份数据库：" + result["path"])
    except Exception as exc:
        print("备份失败：" + str(exc))

    cfg = load_config()
    preferred = int(cfg.get("port", 5000) or 5000)
    port, already = _resolve_port(preferred)
    url = f"http://127.0.0.1:{port}/"
    if already:
        print("工作台已在运行：" + url)
        _open_browser(url)
        return
    print(f"工作台已启动：{url}")
    _open_browser(url)
    app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False, threaded=True)


if __name__ == "__main__":
    main()
