#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
WP 周排班生成器 (wp-weekly-scheduling)

读取 WP 服务单 (Excel)，按自然周（EY 财年：周六 -> 周五）生成：
  Sheet1 周排班-纵表      (唯一数据源，人工只改这里)
  Sheet2 周排班-宽表      (SUMIFS 公式引用纵表)
  Sheet3 周校验           (周标准工时、UT比例(%)、超硬上限标红；intern 以 40h 为基准)
  Sheet4 项目工时校验     (公式，diff>=5h 标黄)
  Sheet5 项目关键节点     (日期 + 提醒列公式)
  Sheet6 日历提醒汇总
  + 项目关键节点.ics
  + run_log.txt

所有读取按表头名称动态定位，人员清单动态读取，支持人员增减替换。
周容量填充按“项目完成优先级”（阶段交付日期升序）；先按目标工时、顺延，末周上浮到硬上限。
"""

import argparse
import calendar
import datetime as dt
import os
import re
import sys
from collections import defaultdict

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# 常量 / 样式
# ---------------------------------------------------------------------------
DEFAULT_PERSONS = [
    ("Marsh", "manager", 0.08, ""),
    ("Vivi", "Senior", 0.22, ""),
    ("Yuan", "staff", 0.7, "公募"),
    ("Luna", "staff", 0.7, "PE"),
    ("Rainbow", "staff", 0.7, "公募"),
    ("Intern", "intern", 0.0, "公募"),
]
MISSING_TOKENS = ("NA", "N/A", "#N/A", "NAN", "NONE", "-")
# 预审工时比例（占 剩余hours）：默认 30%，可按项目类型覆盖（后续人工调整此表即可）
DEFAULT_PRE_RATIO = 0.30
PRE_RATIO_BY_TYPE = {}
FUZZY_DAY = {
    "初": 5, "上旬": 5, "月初": 5,
    "中": 15, "中旬": 15, "月中": 15,
    "底": 25, "下旬": 25, "月底": 25, "末": 25,
}
SHEET_SOURCE = "源数据-WP服务单"
SHEET_BASE = "项目基础"
SHEET_VERT = "周排班-纵表"
SHEET_WIDE = "周排班-宽表"
SHEET_CHECK = "周校验"
SHEET_PROJ = "项目工时校验"
SHEET_NODES = "项目关键节点"
SHEET_CAL = "日历提醒汇总"

HEAD_FILL = PatternFill("solid", fgColor="1F4E78")
HEAD_FONT = Font(name="黑体", bold=True, color="FFFFFF", size=11)
FONT_TEXT = Font(name="黑体", size=11)
FONT_NUM = Font(name="黑体", size=11)
RED_FILL = PatternFill("solid", fgColor="FFC7CE")
YELLOW_FILL = PatternFill("solid", fgColor="FFEB9C")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

STAGE_LABEL = {0: "当天", 3: "3天前", 5: "5天前", 7: "7天前"}
STAGE_PRIORITY = {0: 0, 3: 1, 5: 2, 7: 3}


# ---------------------------------------------------------------------------
# 财年周算法（EY：周六 -> 周五）。已用 SG 排班助手 Date Mapping 校验：
#   FY2027 Week1 = 2026.7.4，Week15 = 2026.10.10。
# ---------------------------------------------------------------------------
def fiscal_year_start(fy):
    """财年 fy 的第一个周六 = 距 (fy-1).7.1 最近的周六。"""
    jul1 = dt.date(fy - 1, 7, 1)
    delta = (jul1.weekday() - 5) % 7
    last_sat = jul1 - dt.timedelta(days=delta)
    next_sat = last_sat + dt.timedelta(days=7)
    if (jul1 - last_sat).days <= (next_sat - jul1).days:
        return last_sat
    return next_sat


def week_start(d):
    """返回 d 所在周六-周五自然周的周六。"""
    return d - dt.timedelta(days=(d.weekday() - 5) % 7)


def week_info(d):
    """返回 (label, week_start_date, fy, week_no)。"""
    ws = week_start(d)
    fy = ws.year + 1
    while fiscal_year_start(fy) > ws:
        fy -= 1
    while fiscal_year_start(fy + 1) <= ws:
        fy += 1
    n = (ws - fiscal_year_start(fy)).days // 7 + 1
    label = "Week{} ({}.{:02d}.{:02d})".format(n, ws.year, ws.month, ws.day)
    return label, ws, fy, n


def is_busy(ws):
    """忙季：1/2/3 月。"""
    return ws.month in (1, 2, 3)


def std_hours(ws):
    """周标准工时：忙季 60h，淡季 40h。"""
    return 60.0 if is_busy(ws) else 40.0


def target_hours(grade, ws):
    """角色周目标：staff=标准，senior=0.85×标准，intern=固定 40h（每周上限）。"""
    s = std_hours(ws)
    if grade == "senior":
        return s * 0.85
    if grade == "intern":
        return 40.0
    return s


def hard_hours(grade, ws):
    """角色周硬上限：标准×1.2；intern 固定 40h/周；manager 不受限（由调用方排除）。"""
    if grade == "intern":
        return 40.0
    return std_hours(ws) * 1.2


def label_to_week(label):
    """从 'Week15 (2026.10.10)' 解析周六日期。"""
    m = re.search(r"\((\d{4})\.(\d{1,2})\.(\d{1,2})\)", str(label))
    if not m:
        return None
    return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))


# ---------------------------------------------------------------------------
# 解析工具
# ---------------------------------------------------------------------------
def norm_name(s):
    s = str(s).replace("\u3000", " ")
    s = re.sub(r"\s+", "", s)
    s = s.replace("（", "(").replace("）", ")")
    return s.lower()


def to_float(v):
    if v is None:
        return 0.0
    try:
        if pd.isna(v):
            return 0.0
    except (TypeError, ValueError):
        pass
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(",", "").strip()
    if s == "" or s.upper() in MISSING_TOKENS:
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def to_text(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def parse_date(v, log=None, desc=""):
    if v is None:
        return None
    if isinstance(v, pd.Timestamp):
        return None if pd.isna(v) else v.date()
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = str(v).strip()
    if s == "" or s.upper() in MISSING_TOKENS:
        return None
    try:
        return pd.to_datetime(s, errors="raise").date()
    except Exception:
        pass
    m = re.search(
        r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(上旬|中旬|下旬|月初|月中|月底|初|中|底|末)?",
        s,
    )
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        mod = m.group(3)
        day = FUZZY_DAY.get(mod, 1) if mod else 1
        try:
            day = min(day, calendar.monthrange(y, mo)[1])
            d = dt.date(y, mo, day)
            if log is not None:
                log.append("  模糊日期近似：{} '{}' -> {}".format(desc, s, d.isoformat()))
            return d
        except ValueError:
            pass
    if log is not None:
        log.append("  无法解析日期：{} '{}'，视为缺失".format(desc, s))
    return None


def person_key(name):
    return re.split(r"[.@\s]", str(name).strip().lower())[0]


def norm_grade(g):
    s = str(g).strip().lower()
    if "man" in s:
        return "manager"
    if "sen" in s:
        return "senior"
    if "int" in s:
        return "intern"
    if "sta" in s:
        return "staff"
    return "staff"


def pre_ratio(proj_type):
    """预审工时占 剩余hours 的比例：按项目类型可覆盖，默认 DEFAULT_PRE_RATIO。"""
    return PRE_RATIO_BY_TYPE.get(str(proj_type).strip(), DEFAULT_PRE_RATIO)


# ---------------------------------------------------------------------------
# 读取 WP 服务单 / 人员清单 / 生成文件回读
# ---------------------------------------------------------------------------
def pick_source_sheet(xls):
    if SHEET_SOURCE in xls.sheet_names:
        return SHEET_SOURCE
    for name in xls.sheet_names:
        if "WP" in name and "服务单" in name:
            return name
    return xls.sheet_names[0]


def load_source(path, log):
    xls = pd.ExcelFile(path)
    sheet = pick_source_sheet(xls)
    df = pd.read_excel(xls, sheet_name=sheet)
    ncols = {}
    for c in df.columns:
        ncols.setdefault(norm_name(c), c)
    return df, ncols


def load_persons(path, log):
    xls = pd.ExcelFile(path)
    sheet = next((n for n in xls.sheet_names if "人员清单" in n), None)
    if sheet is None:
        log.append("未找到『人员清单』sheet，使用内置默认清单。")
        return [{"name": n, "grade": g, "pref": p} for n, g, _w, p in DEFAULT_PERSONS]
    try:
        df = pd.read_excel(xls, sheet_name=sheet)
        ncols = {norm_name(c): c for c in df.columns}
        name_c, grade_c = ncols.get("name"), ncols.get("grade")
        pref_c = ncols.get("优先项目类型")
        if name_c is None or grade_c is None:
            raise ValueError("缺少 Name/Grade 列")
        persons = []
        for _, r in df.iterrows():
            nm = r[name_c]
            if pd.isna(nm) or str(nm).strip() == "":
                continue
            pref = to_text(r[pref_c]) if pref_c is not None else ""
            persons.append({"name": str(nm).strip(), "grade": norm_grade(r[grade_c]), "pref": pref})
        if not persons:
            raise ValueError("清单为空")
        return persons
    except Exception as e:  # noqa: BLE001
        log.append("『人员清单』读取失败({})，使用内置默认清单。".format(e))
        return [{"name": n, "grade": g, "pref": p} for n, g, _w, p in DEFAULT_PERSONS]


def match_person(part, persons, index):
    if part is None:
        return None
    p = str(part).strip()
    if p == "":
        return None
    pl = p.lower()
    if pl in index:
        return index[pl]
    for prs in persons:
        if prs["name"].strip().lower().startswith(pl):
            return prs
    return None


def build_person_index(persons):
    index = {}
    for prs in persons:
        index.setdefault(prs["name"].strip().lower(), prs)
        index.setdefault(person_key(prs["name"]), prs)
    return index


# ---------------------------------------------------------------------------
# 工时分摊
# ---------------------------------------------------------------------------
def distribute_phase(total, start, end):
    """把 total 从 start 铺到 end：先按月份天数拆分，再在月内按自然周天数均分。"""
    res = defaultdict(float)
    if total <= 0 or start is None:
        return res
    if end is None or end < start:
        end = start
    total_days = (end - start).days + 1
    m = dt.date(start.year, start.month, 1)
    while m <= end:
        last = dt.date(m.year, m.month, calendar.monthrange(m.year, m.month)[1])
        seg_s, seg_e = max(start, m), min(end, last)
        if seg_e >= seg_s:
            dim = (seg_e - seg_s).days + 1
            month_hours = total * dim / total_days
            wk = week_start(seg_s)
            while wk <= seg_e:
                ov_s, ov_e = max(seg_s, wk), min(seg_e, wk + dt.timedelta(days=6))
                od = (ov_e - ov_s).days + 1
                if od > 0:
                    res[wk] += month_hours * od / dim
                wk += dt.timedelta(days=7)
        m = dt.date(m.year + 1, 1, 1) if m.month == 12 else dt.date(m.year, m.month + 1, 1)
    return res


def _role_buckets(entry):
    """返回 (manager 占比, senior 占比, staff+intern 池占比)。缺角色时缺口并入池。"""
    managers, seniors = entry["managers"], entry["seniors"]
    staffs, interns = entry["staffs"], entry["interns"]
    mgrsh = 0.08 if managers else 0.0
    srsh = 0.22 if seniors else 0.0
    poolsh = 1.0 - mgrsh - srsh
    if not (staffs or interns):
        if seniors:
            srsh += poolsh
            poolsh = 0.0
        elif managers:
            mgrsh += poolsh
            poolsh = 0.0
        else:
            poolsh = 0.0
    return mgrsh, srsh, poolsh


def allocate(entries, priority, grades, log, reserved=None, prefs=None):
    """全局按周、按项目优先级分配工时。

    entries[eid] 含 code/wp/phase/end/type/week_hours/managers/seniors/staffs/interns。
    规则：manager 8% 均摊不限；senior 22%；staff+intern 共享 70% 池（staff 先到目标、
    剩余给 intern）。每周先填到角色目标，装不下顺延；仅在该条目“末周”允许上浮到硬上限。
    prefs: {name: 优先项目类型}，对 70% 池做软偏好（偏好类型优先于交付日期）。
    返回 (person, week, eid, hours) 列表；项目总工时守恒（diff 不受影响）。
    reserved: {(person, week): hours} 基线已占用（append），计入初始用量。
    """
    reserved = reserved or {}
    prefs = prefs or {}
    used = defaultdict(float)
    for (p, w), h in reserved.items():
        used[(p, w)] += h
    all_weeks = set()
    for e in entries.values():
        all_weeks.update(e["week_hours"].keys())
    if not all_weeks:
        return []
    start, last = min(all_weeks), max(all_weeks)
    span = (last - start).days // 7
    carry_sr = defaultdict(float)
    carry_pool = defaultdict(float)
    results = defaultdict(lambda: defaultdict(float))
    up_logged = set()
    int_logged = set()
    i = 0
    while True:
        w = start + dt.timedelta(days=7 * i)
        row = [eid for eid, e in entries.items()
               if w in e["week_hours"] or carry_sr[eid] > 1e-9 or carry_pool[eid] > 1e-9]
        row.sort(key=lambda x: priority[x])
        busy = is_busy(w)
        pool_rem = {}
        is_last_map = {}
        # 第一遍：manager / senior 分配，并记录各条目本周 pool 需求
        for eid in row:
            e = entries[eid]
            base_w = e["week_hours"].get(w, 0.0)
            managers, seniors = e["managers"], e["seniors"]
            mgrsh, srsh, poolsh = _role_buckets(e)
            is_last = (w == week_start(e["end"]))
            is_last_map[eid] = is_last

            # manager：8% 均摊，不受上限
            if managers and base_w * mgrsh > 1e-9:
                per = base_w * mgrsh / len(managers)
                for m in managers:
                    used[(m, w)] += per
                    results[(m, w)][eid] += per

            # senior：22%
            sr_total = base_w * srsh + carry_sr[eid]
            if seniors and sr_total > 1e-9:
                per = sr_total / len(seniors)
                assigned = 0.0
                for s in seniors:
                    g = grades.get(s, "staff")
                    cap = hard_hours(g, w) if is_last else target_hours(g, w)
                    take = max(0.0, min(per, cap - used[(s, w)]))
                    if take > 1e-9:
                        used[(s, w)] += take
                        results[(s, w)][eid] += take
                        assigned += take
                carry_sr[eid] = max(0.0, sr_total - assigned)
            else:
                carry_sr[eid] = 0.0

            pool_rem[eid] = base_w * poolsh + carry_pool[eid]

        # pool 分配（staff+intern）：按“优先项目类型”软偏好两轮，末周再上浮
        def _pfill(eid, people, cap_fn):
            rem = pool_rem[eid]
            for nm in people:
                if rem <= 1e-9:
                    break
                take = max(0.0, min(rem, cap_fn(nm) - used[(nm, w)]))
                if take > 1e-9:
                    used[(nm, w)] += take
                    results[(nm, w)][eid] += take
                    rem -= take
            pool_rem[eid] = rem

        def _pref_ok(eid, nm):
            p = prefs.get(nm, "")
            return (p == "") or (p == entries[eid].get("type", ""))

        # 第 1 轮（偏好优先）：偏好类型=本项目类型 或 无偏好 的人先到目标
        for eid in row:
            e = entries[eid]
            _pfill(eid, [n for n in e["staffs"] if _pref_ok(eid, n)],
                   lambda n: target_hours("staff", w))
            if busy:
                _pfill(eid, [n for n in e["interns"] if _pref_ok(eid, n)],
                       lambda n: target_hours("intern", w))
        # 第 2 轮（兜底）：其余人到目标
        for eid in row:
            e = entries[eid]
            _pfill(eid, e["staffs"], lambda n: target_hours("staff", w))
            if busy:
                _pfill(eid, e["interns"], lambda n: target_hours("intern", w))
        # 第 3 轮（上浮）：仅末周，staff 优先于 intern 到硬上限
        for eid in row:
            if not is_last_map[eid]:
                continue
            e = entries[eid]
            _pfill(eid, e["staffs"], lambda n: hard_hours("staff", w))
            if busy:
                _pfill(eid, e["interns"], lambda n: hard_hours("intern", w))
        for eid in row:
            carry_pool[eid] = max(0.0, pool_rem[eid])

        # 上浮记录 & intern 无活可排
        people = set()
        for eid in row:
            e = entries[eid]
            people.update(e["managers"] + e["seniors"] + e["staffs"] + e["interns"])
        for p in people:
            g = grades.get(p, "staff")
            if g == "manager":
                continue
            tgt = target_hours(g, w)
            if used[(p, w)] > tgt + 1e-6 and (p, w) not in up_logged:
                up_logged.add((p, w))
                log.append("  上浮：{} {} 目标{}h → 实际{}h（交付周内无法在标准工时消化）".format(
                    p, week_info(w)[0], round(tgt, 2), round(used[(p, w)], 2)))
        if busy:
            for eid in row:
                for inn in entries[eid]["interns"]:
                    if used[(inn, w)] <= 1e-9 and (inn, eid) not in int_logged:
                        int_logged.add((inn, eid))
                        log.append("  intern 无活可排：{} {}（staff 已占用 70% 池）".format(
                            inn, entries[eid]["code"]))

        i += 1
        more = any(v > 1e-9 for v in carry_sr.values()) or any(v > 1e-9 for v in carry_pool.values())
        if w >= last and not more:
            break
        if i > span + 3000:
            break

    out = []
    for (p, w), eidd in results.items():
        for eid, h in eidd.items():
            if h > 1e-9:
                out.append((p, w, eid, h))
    return out


# ---------------------------------------------------------------------------
# 构建排班
# ---------------------------------------------------------------------------
def build_schedule(df, ncols, persons, log, non_interactive, only_keys=None, reserved=None):
    index = build_person_index(persons)
    grades = {p["name"]: p["grade"] for p in persons}
    prefs = {p["name"]: p.get("pref", "") for p in persons}

    def col(*cands):
        for c in cands:
            key = norm_name(c)
            if key in ncols:
                return ncols[key]
        return None

    c_client = col("Client Name")
    c_code = col("Engagement Code")
    c_eng = col("Engagement Name")
    c_order = col("订单编号")
    c_wp = col("WP服务单号")
    c_remain = col("剩余hours")
    c_type = col("项目类型")
    c_pre_s = col("预审开始时间")
    c_pre_e = col("预审交付时间")
    c_ann_s = col("年审开始时间")
    c_ann_e = col("年审交付时间")
    c_report = col("Audit Report Date")
    p_cols = [c for c in df.columns if str(c).startswith("参与人员")]

    meta = {}
    rmeta = {}
    base = {}
    code_name = {}
    qual_pmax = {}
    qual_amin = {}
    report_over = {}
    rows = []
    for _, r in df.iterrows():
        code = to_text(r[c_code]) if c_code else ""
        if code == "":
            continue
        wp = to_text(r[c_wp]) if c_wp else ""
        if only_keys is not None and (code, wp) not in only_keys:
            continue
        rows.append(r)
        if code not in code_name and c_eng:
            code_name[code] = to_text(r[c_eng])
        for cc in (c_pre_s, c_pre_e):
            if cc is None:
                continue
            d = parse_date(r[cc])
            if d and d > qual_pmax.get(code, dt.date.min):
                qual_pmax[code] = d
        for cc in (c_ann_s, c_ann_e):
            if cc is None:
                continue
            d = parse_date(r[cc])
            if d and d < qual_amin.get(code, dt.date.max):
                qual_amin[code] = d
        if c_report:
            rep = parse_date(r[c_report])
            if rep:
                for pname, cc in (("预审", c_pre_e), ("年审", c_ann_e)):
                    if cc is None:
                        continue
                    d = parse_date(r[cc])
                    if d and d > rep and code not in report_over:
                        report_over[code] = (pname, d, rep)

    flagged = {c: (qual_pmax[c], qual_amin[c]) for c in qual_pmax
               if c in qual_amin and qual_pmax[c] > qual_amin[c]}

    skip_codes = set()

    # 数据质量①：预审时间晚于年审时间（任一预审 vs 任一年审）
    if flagged:
        names = sorted(flagged)
        if not non_interactive:
            print("\n以下项目『预审时间晚于年审时间』，疑似数据有误：")
            for code in names[:50]:
                pmax, amin = flagged[code]
                print("  - {} {} 预审({}) > 年审({})".format(code, code_name.get(code, ""), pmax, amin))
            try:
                ans = input("是否跳过这些项目? [Y/n]: ").strip().lower()
            except EOFError:
                ans = "y"
            do_skip = ans in ("", "y", "yes")
        else:
            do_skip = True
        for code in names:
            pmax, amin = flagged[code]
            if do_skip:
                skip_codes.add(code)
                log.append("  跳过项目：{} {} 预审时间({})晚于年审时间({})，疑似数据有误".format(
                    code, code_name.get(code, ""), pmax, amin))
            else:
                log.append("  数据存疑(未跳过)：{} {} 预审时间({})晚于年审时间({})".format(
                    code, code_name.get(code, ""), pmax, amin))

    # 数据质量②：交付日超过报告日(Audit Report Date)
    if report_over:
        names = sorted(report_over)
        if not non_interactive:
            print("\n以下项目『交付日超过报告日(Audit Report Date)』，疑似日期有误：")
            for code in names[:50]:
                pname, d, rep = report_over[code]
                print("  - {} {} {}交付({}) > 报告日({})".format(
                    code, code_name.get(code, ""), pname, d, rep))
            try:
                ans = input("是否跳过这些项目? [Y/n]: ").strip().lower()
            except EOFError:
                ans = "y"
            do_skip = ans in ("", "y", "yes")
        else:
            do_skip = True
        for code in names:
            pname, d, rep = report_over[code]
            if do_skip:
                skip_codes.add(code)
                log.append("  跳过项目：{} {} {}交付({})超过报告日({})，疑似日期有误".format(
                    code, code_name.get(code, ""), pname, d, rep))
            else:
                log.append("  数据存疑(未跳过)：{} {} {}交付({})超过报告日({})，排班截断到报告日所在周".format(
                    code, code_name.get(code, ""), pname, d, rep))

    # 构建条目（含角色名单与周需求），并按报告日截断
    entries = {}
    order_idx = 0
    trunc_seen = set()
    for r in rows:
        code = to_text(r[c_code]) if c_code else ""
        wp = to_text(r[c_wp]) if c_wp else ""
        proj_type = to_text(r[c_type]) if c_type else ""
        if code not in meta:
            meta[code] = {"type": proj_type, "client": to_text(r[c_client]) if c_client else "",
                          "name": to_text(r[c_eng]) if c_eng else ""}
            base[code] = {"type": proj_type, "name": meta[code]["name"], "remain": 0.0}
        base[code]["remain"] += to_float(r[c_remain]) if c_remain else 0.0
        rmeta[(code, wp)] = {"type": proj_type, "client": to_text(r[c_client]) if c_client else "",
                             "name": meta[code]["name"], "order_no": to_text(r[c_order]) if c_order else ""}
        if code in skip_codes:
            continue

        people, seen = [], set()
        for pc in p_cols:
            prs = match_person(r[pc], persons, index)
            if prs and prs["name"] not in seen:
                seen.add(prs["name"])
                people.append(prs)
        if not people:
            continue
        managers = [p["name"] for p in people if p["grade"] == "manager"]
        seniors = [p["name"] for p in people if p["grade"] == "senior"]
        staffs = [p["name"] for p in people if p["grade"] == "staff"]
        interns = [p["name"] for p in people if p["grade"] == "intern"]
        report = parse_date(r[c_report]) if c_report else None

        # 工时预算：不再使用预计桶；预算 = 剩余hours。预审比例按项目类型（默认 30%），其余给年审。
        row_remain = to_float(r[c_remain]) if c_remain else 0.0
        pre_s = parse_date(r[c_pre_s], log, "{} 预审开始".format(code)) if c_pre_s else None
        pre_e = parse_date(r[c_pre_e], log, "{} 预审交付".format(code)) if c_pre_e else None
        ann_s = parse_date(r[c_ann_s], log, "{} 年审开始".format(code)) if c_ann_s else None
        ann_e = parse_date(r[c_ann_e], log, "{} 年审交付".format(code)) if c_ann_e else None

        has_pre = pre_s is not None
        ratio = pre_ratio(proj_type)
        if row_remain <= 0:
            pre_h = ann_h = 0.0
        elif has_pre:
            pre_h = ratio * row_remain
            ann_h = row_remain - pre_h
        else:
            pre_h = 0.0
            ann_h = row_remain

        phases = []
        if pre_h > 0:
            if pre_e is not None and pre_e < pre_s:
                log.append("  跳过预审排班：{} 交付({}) 早于开始({})".format(code, pre_e, pre_s))
            else:
                phases.append(("预审", pre_h, pre_s, pre_e))
        if ann_h > 0:
            if ann_s is None:
                log.append("  跳过年审排班：{} 缺少年审开始时间".format(code))
            elif ann_e is not None and ann_e < ann_s:
                log.append("  跳过年审排班：{} 交付({}) 早于开始({})".format(code, ann_e, ann_s))
            else:
                phases.append(("年审", ann_h, ann_s, ann_e))

        for phase, total, ps, pe in phases:
            end = pe
            if report is not None and pe is not None and pe > report:
                end = report
                if (code, phase) not in trunc_seen:
                    trunc_seen.add((code, phase))
                    log.append("  报告日截断：{} {} 交付({}) 超过报告日({})，排班截断到 {} 所在周".format(
                        code, phase, pe, report, report))
            eid = (code, wp, phase)
            if eid not in entries:
                order_idx += 1
                entries[eid] = {"code": code, "wp": wp, "phase": phase, "deliver": pe,
                                "type": proj_type,
                                "end": end if end is not None else ps, "order": order_idx,
                                "week_hours": defaultdict(float),
                                "managers": [], "seniors": [], "staffs": [], "interns": []}
            e = entries[eid]
            for wk, h in distribute_phase(total, ps, end).items():
                e["week_hours"][wk] += h
            for lst, addl in ((e["managers"], managers), (e["seniors"], seniors),
                              (e["staffs"], staffs), (e["interns"], interns)):
                for nm in addl:
                    if nm not in lst:
                        lst.append(nm)
            if pe is not None and (e["deliver"] is None or pe < e["deliver"]):
                e["deliver"] = pe
            if end is not None and (e["end"] is None or end < e["end"]):
                e["end"] = end

    for e in entries.values():
        if e["end"] is None and e["week_hours"]:
            e["end"] = min(e["week_hours"])

    # 项目完成优先级：阶段交付日期升序 -> Code -> 输入顺序；无交付日期最后
    def prio_key(eid):
        info = entries[eid]
        d = info["deliver"]
        return (d is None, d or dt.date.min, info["code"], info["order"])

    priority = {eid: prio_key(eid) for eid in entries}
    if entries:
        log.append("项目优先级排序（按阶段交付日期，早者优先）：")
        for idx, eid in enumerate(sorted(entries, key=lambda e: priority[e]), 1):
            info = entries[eid]
            dd = info["deliver"].isoformat() if info["deliver"] else "无交付日期(最后处理)"
            log.append("  {}. {} {} {}".format(idx, info["code"], info["phase"], dd))

    # 分配
    assignments = allocate(entries, priority, grades, log, reserved, prefs)
    max_week_by_entry = {}
    for p, wk, eid, h in assignments:
        if h <= 1e-9:
            continue
        if eid not in max_week_by_entry or wk > max_week_by_entry[eid]:
            max_week_by_entry[eid] = wk

    records = []
    for p, wk, eid, h in assignments:
        h2 = round(h, 2)
        if h2 <= 0:
            continue
        e = entries[eid]
        label, _, _, _ = week_info(wk)
        info = rmeta.get((e["code"], e["wp"]), {})
        records.append({
            "type": info.get("type", ""),
            "client": info.get("client", ""),
            "code": e["code"],
            "name": info.get("name", ""),
            "order": info.get("order_no", ""),
            "wp": e["wp"],
            "staff": p,
            "week": wk,
            "label": label,
            "hours": h2,
        })

    # 剩余容量不足 / 容量冲突（因优先级低被顺延到交付周之后）
    n_conflict = 0
    for eid in sorted(entries, key=lambda e: priority[e]):
        info = entries[eid]
        d = info["deliver"]
        mw = max_week_by_entry.get(eid)
        if d is None or mw is None:
            continue
        if mw > week_start(d):
            n_conflict += 1
            log.append("  因优先级低，剩余容量不足：{} {} 已排到 {}（交付周 {}）".format(
                info["code"], info["phase"], week_info(mw)[0], week_info(week_start(d))[0]))
    if reserved and n_conflict:
        log.append("  容量冲突汇总：{} 处新增工时因基线占用被顺延至交付周之后".format(n_conflict))

    return records, meta, base, skip_codes


def read_existing_vertical(path):
    wb = load_workbook(path, data_only=True)
    if SHEET_VERT not in wb.sheetnames:
        return None
    ws = wb[SHEET_VERT]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return None
    records = []
    for r in rows[1:]:
        if r[0] is None and r[2] is None:
            continue
        code = to_text(r[2])
        label = to_text(r[7])
        wk = label_to_week(label)
        if wk is None:
            continue
        records.append({
            "type": to_text(r[0]), "client": to_text(r[1]), "code": code,
            "name": to_text(r[3]), "order": to_text(r[4]), "wp": to_text(r[5]),
            "staff": to_text(r[6]), "week": wk, "label": label,
            "hours": to_float(r[8]),
        })
    return records or None


def read_existing_base(path):
    wb = load_workbook(path, data_only=True)
    if SHEET_BASE not in wb.sheetnames:
        return None
    ws = wb[SHEET_BASE]
    out = {}
    for r in list(ws.iter_rows(values_only=True))[1:]:
        if r[0] is None:
            continue
        out[to_text(r[0])] = {
            "name": to_text(r[1]), "type": to_text(r[2]), "remain": to_float(r[3]),
        }
    return out or None


def read_existing_nodes(path):
    wb = load_workbook(path, data_only=True)
    if SHEET_NODES not in wb.sheetnames:
        return None
    ws = wb[SHEET_NODES]
    edited = {}
    for r in list(ws.iter_rows(values_only=True))[1:]:
        if r[0] is None or r[2] is None:
            continue
        edited[(to_text(r[0]), to_text(r[2]))] = parse_date(r[3])
    return edited or None


def read_existing_node_rows(path):
    """读取基线『项目关键节点』的完整行（保留人工修改的日期与参与人员）。"""
    wb = load_workbook(path, data_only=True)
    if SHEET_NODES not in wb.sheetnames:
        return []
    ws = wb[SHEET_NODES]
    rows = []
    for r in list(ws.iter_rows(values_only=True))[1:]:
        if r[0] is None or r[2] is None:
            continue
        d = parse_date(r[3])
        if d is None:
            continue
        rows.append({"code": to_text(r[0]), "name": to_text(r[1]),
                     "type": to_text(r[2]), "date": d, "people": to_text(r[4])})
    rows.sort(key=lambda x: (x["date"], x["code"], x["type"]))
    return rows


def find_latest_base(output_dir, log):
    """在 output_dir 下找最新的、含『周排班-纵表』的 WP周排班_*.xlsx。"""
    if not os.path.isdir(output_dir):
        return None
    cands = []
    for f in os.listdir(output_dir):
        if not (f.startswith("WP周排班") and f.lower().endswith(".xlsx")) or f.startswith("~$"):
            continue
        p = os.path.join(output_dir, f)
        try:
            wb = load_workbook(p, read_only=True, data_only=True)
            ok = SHEET_VERT in wb.sheetnames
            wb.close()
        except Exception:  # noqa: BLE001
            ok = False
        if ok:
            cands.append(p)
    if not cands:
        return None
    return max(cands, key=os.path.getmtime)


def merge_base(a, b):
    """合并项目基础信息（按 Engagement Code），remain 相加。"""
    out = {k: dict(v) for k, v in (a or {}).items()}
    for k, v in (b or {}).items():
        if k in out:
            out[k]["remain"] = round(out[k].get("remain", 0.0) + v.get("remain", 0.0), 2)
            if not out[k].get("name"):
                out[k]["name"] = v.get("name", "")
            if not out[k].get("type"):
                out[k]["type"] = v.get("type", "")
        else:
            out[k] = dict(v)
    return out


def build_nodes(df, ncols, persons, log, edited=None, skip_codes=None, only_keys=None):
    skip_codes = skip_codes or set()
    index = build_person_index(persons)
    c_code = next((ncols[norm_name(c)] for c in ["Engagement Code"] if norm_name(c) in ncols), None)
    c_eng = next((ncols[norm_name(c)] for c in ["Engagement Name"] if norm_name(c) in ncols), None)
    c_wp = next((ncols[norm_name(c)] for c in ["WP服务单号"] if norm_name(c) in ncols), None)
    date_cols = [
        ("预审开始", "预审开始时间"), ("预审交付", "预审交付时间"),
        ("年审开始", "年审开始时间"), ("年审交付", "年审交付时间"),
    ]
    p_cols = [c for c in df.columns if str(c).startswith("参与人员")]
    nodes = {}
    for _, r in df.iterrows():
        code = to_text(r[c_code]) if c_code else ""
        if code == "" or code in skip_codes:
            continue
        if only_keys is not None and (code, to_text(r[c_wp]) if c_wp else "") not in only_keys:
            continue
        names, seen = [], set()
        for pc in p_cols:
            prs = match_person(r[pc], persons, index)
            if prs and prs["name"] not in seen:
                seen.add(prs["name"])
                names.append(prs["name"])
        people = "、".join(names)
        for ntype, cname in date_cols:
            key = norm_name(cname)
            if key not in ncols:
                continue
            d = parse_date(r[ncols[key]], log, "{} {}".format(code, ntype))
            if d is None:
                continue
            nodes[(code, ntype)] = {
                "code": code,
                "name": to_text(r[c_eng]) if c_eng else "",
                "type": ntype,
                "date": d,
                "people": people,
            }
    if edited:
        for key, val in edited.items():
            if val is None:
                nodes.pop(key, None)
            elif key in nodes:
                nodes[key]["date"] = val
    out = [n for n in nodes.values() if n["date"] is not None]
    out.sort(key=lambda x: (x["date"], x["code"], x["type"]))
    return out


# ---------------------------------------------------------------------------
# Excel 写入
# ---------------------------------------------------------------------------
def is_numeric_cell(v):
    """数字/日期/公式（结果多为数字）→ True，用于选择微软雅黑字体。"""
    if v is None or isinstance(v, bool):
        return False
    if isinstance(v, (int, float)):
        return True
    if isinstance(v, (dt.date, dt.datetime)):
        return True
    if isinstance(v, str) and v.startswith("="):
        return True
    return False


def style_header(ws, ncols, row=1):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEAD_FILL
        cell.font = HEAD_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def _disp_width(s):
    return sum(2 if ord(ch) > 127 else 1 for ch in str(s))


def fit_header_height(ws, ncols, row=1):
    """按表头文本宽度与列宽估算换行行数，自动设置表头行高（最小 18）。"""
    max_lines = 1
    for c in range(1, ncols + 1):
        text = ws.cell(row=row, column=c).value
        if text is None:
            continue
        w = ws.column_dimensions[get_column_letter(c)].width or 8.43
        per = max(int(w) - 1, 1)
        lines = (_disp_width(text) + per - 1) // per
        max_lines = max(max_lines, max(lines, 1))
    ws.row_dimensions[row].height = min(15 * max_lines + 3, 90)


def style_body(ws, ncols, header_row=1):
    """表体统一：黑体、细边框、垂直居中。"""
    for row in ws.iter_rows(min_row=header_row + 1):
        for cell in row:
            if cell.column > ncols:
                continue
            cell.font = FONT_NUM if is_numeric_cell(cell.value) else FONT_TEXT
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center")


def autosize(ws, max_width=42, min_width=9):
    """按表头与非公式内容测算列宽（忽略公式字符串）。"""
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        width = min_width
        for cell in col:
            v = cell.value
            if v is None:
                continue
            if isinstance(v, str) and v.startswith("="):
                continue
            width = max(width, min(max_width, len(str(v)) + 2))
        ws.column_dimensions[letter].width = width


def write_vertical(wb, records):
    ws = wb.create_sheet(SHEET_VERT)
    headers = ["项目类型", "Client Name", "Engagement Code", "Engagement Name",
               "订单编号", "WP服务单号", "Staff Name", "周期", "hours"]
    ws.append(headers)
    records = sorted(records, key=lambda x: (x["week"], x["order"], x["staff"], x["code"]))
    for rec in records:
        ws.append([rec["type"], rec["client"], rec["code"], rec["name"],
                   rec["order"], rec["wp"], rec["staff"], rec["label"], rec["hours"]])
    style_header(ws, len(headers))
    style_body(ws, len(headers))
    ws.freeze_panes = "A2"
    autosize(ws)
    fit_header_height(ws, len(headers))
    return records


def write_wide(wb, records, meta):
    ws = wb.create_sheet(SHEET_WIDE)
    weeks = sorted({r["week"] for r in records})
    if weeks:
        full, w = [], weeks[0]
        while w <= weeks[-1]:
            full.append(w)
            w += dt.timedelta(days=7)
        weeks = full
    headers = ["项目类型", "Client Name", "Engagement Code", "Engagement Name",
               "订单编号", "WP服务单号", "Staff Name"]
    headers += [week_info(w)[0] for w in weeks]
    ws.append(headers)
    pairs = {}
    for r in records:
        pairs.setdefault((r["code"], r["staff"]), r)
    ordered = sorted(pairs.items(), key=lambda kv: (kv[1]["type"], kv[1]["code"], kv[1]["staff"]))
    for i, (_key, r) in enumerate(ordered, start=2):
        ws.append([r["type"], r["client"], r["code"], r["name"], r["order"], r["wp"], r["staff"]])
        for j in range(len(weeks)):
            col = get_column_letter(8 + j)
            ws.cell(row=i, column=8 + j).value = (
                "=SUMIFS('{v}'!$I:$I,'{v}'!$C:$C,$C{r},"
                "'{v}'!$G:$G,$G{r},'{v}'!$H:$H,{c}$1)".format(
                    v=SHEET_VERT, r=i, c=col
                )
            )
    style_header(ws, len(headers))
    style_body(ws, len(headers))
    ws.freeze_panes = "H2"
    autosize(ws)
    for j in range(len(weeks)):
        ws.column_dimensions[get_column_letter(8 + j)].width = 11
    fit_header_height(ws, len(headers))


def write_check(wb, records):
    ws = wb.create_sheet(SHEET_CHECK)
    headers = ["Staff Name", "周期", "该周合计 hours", "是否忙季", "周标准工时",
               "UT比例", "是否超硬上限"]
    ws.append(headers)
    rows = sorted({(r["staff"], r["week"], r["label"]) for r in records},
                  key=lambda x: (x[1], x[0]))
    for i, (staff, _wk, label) in enumerate(rows, start=2):
        ws.cell(row=i, column=1, value=staff)
        ws.cell(row=i, column=2, value=label)
        ws.cell(row=i, column=3,
                value="=SUMIFS('{}'!$I:$I,'{}'!$G:$G,$A{r},'{}'!$H:$H,$B{r})".format(SHEET_VERT, SHEET_VERT, SHEET_VERT, r=i))
        mm = "MID($B{r},FIND(\"(\",$B{r})+6,2)".format(r=i)
        ws.cell(row=i, column=4,
                value='=IF(OR({m}="01",{m}="02",{m}="03"),"是","否")'.format(m=mm))
        # 角色（intern?）来自隐藏『人员清单』sheet
        g = 'IFERROR(VLOOKUP($A{r},\'人员清单\'!$A:$B,2,FALSE),"")'.format(r=i)
        # 周标准工时：intern=40；其余淡季 40 / 忙季 60
        ws.cell(row=i, column=5,
                value='=IF({g}="intern",40,IF($D{r}="是",60,40))'.format(g=g, r=i))
        # UT比例：合计 / 40（以 40h 为基准），百分比格式
        ut = ws.cell(row=i, column=6,
                     value='=ROUND($C{r}/40,4)'.format(r=i))
        ut.number_format = "0%"
        # 是否超硬上限：intern>40；非 intern 忙季>78 / 淡季>48
        ws.cell(row=i, column=7,
                value='=IF({g}="intern",IF($C{r}>40,"是","否"),'
                      'IF($D{r}="是",IF($C{r}>78,"是","否"),IF($C{r}>48,"是","否")))'.format(g=g, r=i))
    style_header(ws, len(headers))
    style_body(ws, len(headers))
    if rows:
        ws.conditional_formatting.add(
            "G2:G{}".format(len(rows) + 1),
            FormulaRule(formula=['$G2="是"'], fill=RED_FILL),
        )
    ws.freeze_panes = "A2"
    autosize(ws)
    fit_header_height(ws, len(headers))


def write_proj(wb, records, base):
    ws = wb.create_sheet(SHEET_PROJ)
    headers = ["Engagement Code", "Engagement Name", "剩余hours",
               "排班合计 hours", "diff", "diff 是否 <5h"]
    ws.append(headers)
    codes = sorted({r["code"] for r in records} | set(base.keys()))
    for i, code in enumerate(codes, start=2):
        info = base.get(code, {})
        ws.cell(row=i, column=1, value=code)
        ws.cell(row=i, column=2, value=info.get("name", ""))
        ws.cell(row=i, column=3, value=round(info.get("remain", 0.0), 2))
        ws.cell(row=i, column=4,
                value="=SUMIFS('{}'!$I:$I,'{}'!$C:$C,$A{r})".format(SHEET_VERT, SHEET_VERT, r=i))
        ws.cell(row=i, column=5, value="=$D{r}-$C{r}".format(r=i))
        ws.cell(row=i, column=6, value='=IF(ABS($E{r})<5,"是","否")'.format(r=i))
    style_header(ws, len(headers))
    style_body(ws, len(headers))
    if codes:
        ws.conditional_formatting.add(
            "F2:F{}".format(len(codes) + 1),
            FormulaRule(formula=['$F2="否"'], fill=YELLOW_FILL),
        )
    ws.freeze_panes = "A2"
    autosize(ws)
    fit_header_height(ws, len(headers))


def write_nodes(wb, nodes):
    ws = wb.create_sheet(SHEET_NODES)
    headers = ["Engagement Code", "Engagement Name", "节点类型", "日期",
               "参与人员", "提醒7天前", "提醒5天前", "提醒3天前", "提醒当天"]
    ws.append(headers)
    for i, n in enumerate(nodes, start=2):
        ws.cell(row=i, column=1, value=n["code"])
        ws.cell(row=i, column=2, value=n["name"])
        ws.cell(row=i, column=3, value=n["type"])
        dc = ws.cell(row=i, column=4, value=n["date"])
        dc.number_format = "yyyy-mm-dd"
        ws.cell(row=i, column=5, value=n["people"])
        for off, col in [(7, 6), (5, 7), (3, 8), (0, 9)]:
            cell = ws.cell(row=i, column=col, value="=$D{r}-{o}".format(r=i, o=off))
            cell.number_format = "yyyy-mm-dd"
    style_header(ws, len(headers))
    style_body(ws, len(headers))
    ws.freeze_panes = "A2"
    autosize(ws)
    fit_header_height(ws, len(headers))


def build_reminders(nodes):
    groups = defaultdict(list)
    for n in nodes:
        for off in (7, 5, 3, 0):
            rdate = n["date"] - dt.timedelta(days=off)
            groups[rdate].append({
                "node": n,
                "stage": off,
                "days": off,
            })
    rows = []
    for rdate, items in groups.items():
        stages = {it["stage"] for it in items}
        stage = items[0]["stage"] if len(stages) == 1 else None
        stage_label = STAGE_LABEL[stage] if stage is not None else "混合"
        items.sort(key=lambda x: (x["node"]["name"], x["node"]["code"], x["node"]["type"]))
        lines = []
        for it in items:
            n = it["node"]
            lines.append("{}（{}）| {} | {} | 距DDL{}天".format(
                n["name"] or n["code"], n["code"], n["type"], n["date"].isoformat(), it["days"]))
        rows.append({
            "date": rdate,
            "stage": stage,
            "stage_label": stage_label,
            "count": len(items),
            "title": "DDL汇总：{}（{}项）".format(rdate.isoformat(), len(items)),
            "detail": "\n".join(lines),
            "items": items,
        })
    rows.sort(key=lambda x: (
        STAGE_PRIORITY.get(x["stage"], 9) if x["stage"] is not None else 4,
        x["items"][0]["node"]["name"],
    ))
    return rows


def write_calendar(wb, reminders):
    ws = wb.create_sheet(SHEET_CAL)
    headers = ["提醒日期", "提醒阶段", "涉及节点数", "汇总标题", "汇总详情"]
    ws.append(headers)
    for i, row in enumerate(reminders, start=2):
        dc = ws.cell(row=i, column=1, value=row["date"])
        dc.number_format = "yyyy-mm-dd"
        ws.cell(row=i, column=2, value=row["stage_label"])
        ws.cell(row=i, column=3, value=row["count"])
        ws.cell(row=i, column=4, value=row["title"])
        dc2 = ws.cell(row=i, column=5, value=row["detail"])
        dc2.alignment = Alignment(wrap_text=True, vertical="top")
    style_header(ws, len(headers))
    style_body(ws, len(headers))
    ws.freeze_panes = "A2"
    autosize(ws)
    fit_header_height(ws, len(headers))


def write_helpers(wb, df, base, persons):
    ws = wb.create_sheet(SHEET_BASE)
    ws.append(["Engagement Code", "Engagement Name", "项目类型", "剩余hours"])
    for code, info in sorted(base.items()):
        ws.append([code, info.get("name", ""), info.get("type", ""), round(info.get("remain", 0.0), 2)])
    style_header(ws, 4)
    style_body(ws, 4)
    autosize(ws)
    fit_header_height(ws, 4)
    ws.sheet_state = "hidden"

    ws2 = wb.create_sheet(SHEET_SOURCE)
    ws2.append(list(df.columns))
    for _, r in df.iterrows():
        ws2.append([None if pd.isna(v) else v for v in r.tolist()])
    style_header(ws2, len(df.columns))
    style_body(ws2, len(df.columns))
    autosize(ws2)
    fit_header_height(ws2, len(df.columns))
    ws2.sheet_state = "hidden"

    ws3 = wb.create_sheet("人员清单")
    ws3.append(["Name", "Grade", "hours占比", "优先项目类型"])
    for prs in persons:
        ws3.append([prs["name"], prs["grade"], None, prs.get("pref", "")])
    style_header(ws3, 4)
    style_body(ws3, 4)
    autosize(ws3)
    fit_header_height(ws3, 4)
    ws3.sheet_state = "hidden"


def write_ics(path, reminders):
    def esc(s):
        return (str(s).replace("\\", "\\\\").replace(";", "\\;")
                .replace(",", "\\,").replace("\n", "\\n"))

    lines = ["BEGIN:VCALENDAR", "VERSION:2.0",
             "PRODID:-//wp-weekly-scheduling//CN", "CALSCALE:GREGORIAN", "METHOD:PUBLISH"]
    stamp = dt.datetime.now().strftime("%Y%m%dT%H%M%S")
    for idx, row in enumerate(reminders):
        d = row["date"]
        start = d.strftime("%Y%m%d") + "T100000"
        end = d.strftime("%Y%m%d") + "T103000"
        if row["count"] == 1 and row["stage"] == 0:
            n = row["items"][0]["node"]
            title = "DDL提醒：{}{}（今天）".format(n["name"] or n["code"], n["type"])
        else:
            title = row["title"]
        lines += [
            "BEGIN:VEVENT",
            "UID:wp-{}-{}@wp-weekly-scheduling".format(d.strftime("%Y%m%d"), idx),
            "DTSTAMP:{}".format(stamp),
            "DTSTART:{}".format(start),
            "DTEND:{}".format(end),
            "SUMMARY:{}".format(esc(title)),
            "DESCRIPTION:{}".format(esc(row["detail"])),
            "BEGIN:VALARM", "TRIGGER:PT0M", "ACTION:DISPLAY",
            "DESCRIPTION:{}".format(esc(title)),
            "END:VALARM", "END:VEVENT",
        ]
    lines.append("END:VCALENDAR")
    with open(path, "w", encoding="utf-8", newline="") as f:
        f.write("\r\n".join(lines) + "\r\n")


# ---------------------------------------------------------------------------
# Outlook 日历写入 / 本地弹窗检查脚本
# ---------------------------------------------------------------------------
DAILY_CHECK_TEMPLATE = '''# -*- coding: utf-8 -*-
"""WP DDL 每日弹窗检查（由 generate_schedule.py 自动生成）。"""
import argparse
import datetime as dt
import glob
import os
import subprocess

DEFAULT_OUTPUT_DIR = r"__OUTDIR__"


def find_latest(output_dir):
    files = [p for p in glob.glob(os.path.join(output_dir, "WP周排班_*.xlsx"))
             if not os.path.basename(p).startswith("~$")]
    return max(files, key=os.path.getmtime) if files else None


def load_reminders(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    if "日历提醒汇总" not in wb.sheetnames:
        return []
    ws = wb["日历提醒汇总"]
    rows = []
    for r in list(ws.iter_rows(values_only=True))[1:]:
        if r[0] is None:
            continue
        d = r[0]
        if isinstance(d, dt.datetime):
            d = d.date()
        elif isinstance(d, str):
            try:
                d = dt.datetime.strptime(d[:10], "%Y-%m-%d").date()
            except ValueError:
                continue
        else:
            continue
        rows.append({"date": d, "title": r[3] or "", "detail": r[4] or ""})
    return rows


def notify(title, msg):
    try:
        from plyer import notification
        notification.notify(title=title, message=msg, timeout=15)
        return
    except Exception:
        pass
    try:
        from win10toast import ToastNotifier
        ToastNotifier().show_toast(title, msg, duration=15, threaded=False)
        return
    except Exception:
        pass
    ps = ("Add-Type -AssemblyName System.Windows.Forms; "
          "[System.Windows.Forms.MessageBox]::Show(@'\\n" + msg + "\\n'@, '" + title + "')")
    try:
        subprocess.run(["powershell", "-NoProfile", "-Command", ps], check=False)
    except Exception:
        print(title)
        print(msg)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--output-dir", default=DEFAULT_OUTPUT_DIR)
    ap.add_argument("--days", type=int, default=7)
    args = ap.parse_args()
    latest = find_latest(args.output_dir)
    if not latest:
        return
    today = dt.date.today()
    end = today + dt.timedelta(days=args.days)
    rows = [r for r in load_reminders(latest) if today <= r["date"] <= end]
    if not rows:
        return
    rows.sort(key=lambda r: r["date"])
    lines = []
    for r in rows:
        lines.append("{}  {}".format(r["date"].isoformat(), r["title"]))
        if r["detail"]:
            lines.append(r["detail"])
    notify("WP DDL 提醒", "\\n".join(lines))


if __name__ == "__main__":
    main()
'''


def write_daily_scripts(scripts_dir, out_dir):
    """生成 daily_check.py / daily_check.bat / register_task.bat / unregister_task.bat。

    py 程序由 bat 调用；任务计划注册的是 daily_check.bat（每天 10:30）。
    """
    pyexe = sys.executable
    check = os.path.join(scripts_dir, "daily_check.py")
    text = DAILY_CHECK_TEMPLATE.replace("__OUTDIR__", out_dir)
    with open(check, "w", encoding="utf-8", newline="") as f:
        f.write(text.replace("\n", "\r\n") if os.name == "nt" else text)

    # 由 bat 直接调用 python（任务计划执行本 bat）
    launcher = os.path.join(scripts_dir, "daily_check.bat")
    launcher_lines = [
        "@echo off",
        'cd /d "%~dp0."',
        '"{}" "%~dp0daily_check.py"'.format(pyexe),
    ]
    with open(launcher, "w", encoding="gbk", newline="") as f:
        f.write("\r\n".join(launcher_lines) + "\r\n")

    reg = os.path.join(scripts_dir, "register_task.bat")
    reg_lines = [
        "@echo off",
        "cd /d \"%~dp0.\"",
        'schtasks /Create /TN WP_DDL_DailyCheck /TR "\\"%~dp0daily_check.bat\\"" '
        "/SC DAILY /ST 10:30 /F",
        "echo 已注册任务 WP_DDL_DailyCheck（每天 10:30 运行 daily_check.bat）",
        "pause",
    ]
    with open(reg, "w", encoding="gbk", newline="") as f:
        f.write("\r\n".join(reg_lines) + "\r\n")

    unreg = os.path.join(scripts_dir, "unregister_task.bat")
    unreg_lines = [
        "@echo off",
        "schtasks /Delete /TN WP_DDL_DailyCheck /F",
        "echo 已删除任务 WP_DDL_DailyCheck",
        "pause",
    ]
    with open(unreg, "w", encoding="gbk", newline="") as f:
        f.write("\r\n".join(unreg_lines) + "\r\n")
    return [check, launcher, reg, unreg]


def sync_outlook(reminders, log):
    """先删所有带 WP-DDL 标识的事件，再按最新提醒写入。成功返回 True。"""
    try:
        import win32com.client  # noqa: PLC0415
    except Exception as e:  # noqa: BLE001
        log.append("  未安装 pywin32（win32com）：{}".format(e))
        return False
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
        cal = ns.GetDefaultFolder(9)  # olFolderCalendar
        items = cal.Items
        to_del = []
        for it in items:
            try:
                cats = str(getattr(it, "Categories", "") or "")
            except Exception:  # noqa: BLE001
                cats = ""
            if "WP-DDL" in cats:
                to_del.append(it)
        deleted = 0
        for it in to_del:
            try:
                it.Delete()
                deleted += 1
            except Exception:  # noqa: BLE001
                pass
        log.append("  Outlook：已删除旧 WP-DDL 事件 {} 个".format(deleted))

        written = 0
        for row in reminders:
            d = row["date"]
            appt = outlook.CreateItem(1)  # olAppointmentItem
            appt.Subject = row["title"]
            appt.Start = dt.datetime(d.year, d.month, d.day, 10, 0)
            appt.End = dt.datetime(d.year, d.month, d.day, 10, 30)
            appt.Body = row["detail"]
            appt.ReminderSet = True
            appt.ReminderMinutesBeforeStart = 0
            appt.Categories = "WP-DDL,WP-DDL-{}-{}".format(d.isoformat(), row["stage_label"])
            appt.Save()
            written += 1
        log.append("  Outlook：已写入 {} 个 WP-DDL 事件".format(written))
        return True
    except Exception as e:  # noqa: BLE001
        log.append("  Outlook COM 调用失败：{}".format(e))
        return False


def handle_reminders(reminders, args, log, out_dir):
    """按 --remind-mode 输出提醒。返回是否生成了 ics。"""
    mode = getattr(args, "remind_mode", "both")
    log.append("提醒模式：{}".format(mode))
    out_dir = os.path.abspath(out_dir)
    out_ics = os.path.join(out_dir, "项目关键节点.ics")
    scripts_dir = os.path.dirname(os.path.abspath(__file__))
    wrote_ics = False

    if mode == "ics":
        write_ics(out_ics, reminders)
        wrote_ics = True
        log.append("  已生成日历文件：{}".format(out_ics))
    elif mode == "popup":
        files = write_daily_scripts(scripts_dir, out_dir)
        log.append("  已生成弹窗脚本：{}".format("、".join(os.path.basename(f) for f in files)))
    elif mode in ("outlook", "both"):
        ok = sync_outlook(reminders, log)
        if mode == "both" or not ok:
            if not ok:
                log.append("  Outlook 不可用/失败，降级为 ics 模式。")
            write_ics(out_ics, reminders)
            wrote_ics = True
            log.append("  已生成日历文件：{}".format(out_ics))
        if mode == "both":
            files = write_daily_scripts(scripts_dir, out_dir)
            log.append("  已生成弹窗脚本：{}".format("、".join(os.path.basename(f) for f in files)))
    return wrote_ics


# ---------------------------------------------------------------------------
# 输入文件动态扫描
# ---------------------------------------------------------------------------
def scan_candidates(input_dir, log):
    """扫描 input_dir 下所有 .xlsx，返回符合 WP服务单 格式的候选文件列表。

    候选条件（同时满足）：
      1. 含名为『WP服务单』的 sheet（兼容已生成文件的『源数据-WP服务单』）；
      2. 该 sheet 表头同时含『Engagement Code』与『剩余hours』。
    """
    if not os.path.isdir(input_dir):
        return []
    files = sorted(
        os.path.join(input_dir, f)
        for f in os.listdir(input_dir)
        if f.lower().endswith(".xlsx") and not f.startswith("~$")
    )
    candidates = []
    for path in files:
        try:
            xls = pd.ExcelFile(path)
            sheet = next(
                (n for n in xls.sheet_names if "WP服务单" in str(n).replace(" ", "")),
                None,
            )
            if sheet is None:
                continue
            cols = {norm_name(c) for c in pd.read_excel(xls, sheet_name=sheet, nrows=1).columns}
            if norm_name("Engagement Code") in cols and norm_name("剩余hours") in cols:
                candidates.append(path)
        except Exception as e:  # noqa: BLE001
            log.append("  跳过无法读取的文件 {}：{}".format(os.path.basename(path), e))
    return candidates


def select_input(args, log):
    """决定使用哪个输入文件：--input 优先，否则扫描 --input-dir。"""
    if args.input:
        if not os.path.exists(args.input):
            print("错误：指定的 --input 文件不存在：{}".format(args.input))
            sys.exit(1)
        return args.input

    candidates = scan_candidates(args.input_dir, log)
    if not candidates:
        all_xlsx = []
        if os.path.isdir(args.input_dir):
            all_xlsx = sorted(f for f in os.listdir(args.input_dir) if f.lower().endswith(".xlsx"))
        print("未在 {} 目录找到符合 WP服务单 格式的 xlsx 文件，"
              "请检查文件是否包含 WP服务单 sheet 和 Engagement Code / 剩余hours 表头。".format(args.input_dir))
        if all_xlsx:
            print("该目录下的 xlsx 文件：")
            for f in all_xlsx:
                print("  - {}".format(f))
        else:
            print("该目录下没有任何 .xlsx 文件。")
        sys.exit(1)

    if len(candidates) == 1:
        log.append("扫描到唯一候选输入文件：{}".format(os.path.basename(candidates[0])))
        return candidates[0]

    if args.non_interactive:
        chosen = max(candidates, key=os.path.getmtime)
        log.append("扫描到 {} 个候选文件，非交互模式选择修改时间最新的：{}".format(
            len(candidates), os.path.basename(chosen)))
        return chosen

    print("扫描到多个候选输入文件：")
    for i, p in enumerate(candidates, 1):
        mtime = dt.datetime.fromtimestamp(os.path.getmtime(p)).strftime("%Y-%m-%d %H:%M")
        print("  {}. {}  (修改时间 {})".format(i, os.path.basename(p), mtime))
    while True:
        try:
            ans = input("请输入序号选择 (1-{}): ".format(len(candidates))).strip()
        except EOFError:
            ans = "1"
        if ans.isdigit() and 1 <= int(ans) <= len(candidates):
            chosen = candidates[int(ans) - 1]
            log.append("用户选择输入文件：{}".format(os.path.basename(chosen)))
            return chosen
        print("输入无效，请输入 1-{} 之间的序号。".format(len(candidates)))


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def emit(records, base, persons, df, nodes, reminders, input_path, args, log,
         mode_label, base_path=None):
    """校验并写出所有产物（full / append 共用）。"""
    grades = {p["name"]: p["grade"] for p in persons}

    # ---- 校验：超硬上限 / diff ----
    over = defaultdict(float)
    for r in records:
        over[(r["staff"], r["week"])] += r["hours"]
    for (staff, wk), tot in sorted(over.items()):
        g = grades.get(staff, "staff")
        if g == "manager":
            continue
        if tot > hard_hours(g, wk) + 1e-6:
            log.append("  周超硬上限警告：{} {} 合计 {}h > 硬上限 {}h".format(
                staff, week_info(wk)[0], round(tot, 2), round(hard_hours(g, wk), 2)))
    proj_total = defaultdict(float)
    for r in records:
        proj_total[r["code"]] += r["hours"]
    for code, tot in sorted(proj_total.items()):
        diff = tot - base.get(code, {}).get("remain", 0.0)
        if abs(diff) >= 5:
            log.append("  工时偏差：{} 排班{}h 剩余{}h diff={}h".format(
                code, round(tot, 2), base.get(code, {}).get("remain", 0.0), round(diff, 2)))

    # ---- 写出 ----
    wb = Workbook()
    wb.remove(wb.active)
    write_vertical(wb, records)
    write_wide(wb, records, base)
    write_check(wb, records)
    write_proj(wb, records, base)
    write_nodes(wb, nodes)
    write_calendar(wb, reminders)
    write_helpers(wb, df, base, persons)

    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M")
    out_xlsx = os.path.join(args.output_dir, "WP周排班_{}.xlsx".format(stamp))
    wb.save(out_xlsx)
    wrote_ics = handle_reminders(reminders, args, log, args.output_dir)
    out_ics = os.path.join(args.output_dir, "项目关键节点.ics")

    header = ["WP 周排班运行日志",
              "运行时间：{}".format(dt.datetime.now().isoformat(timespec="seconds")),
              "模式：{}".format(mode_label)]
    if base_path:
        header.append("基线文件：{}".format(os.path.abspath(base_path)))
    log_lines = header + [
        "",
        "输入文件：{}".format(os.path.abspath(input_path)),
        "输出文件：{}".format(os.path.abspath(out_xlsx)),
        "日历文件：{}".format(os.path.abspath(out_ics) if wrote_ics else "（未生成）"),
        "排班记录数：{}".format(len(records)),
        "节点数：{}".format(len(nodes)),
        "提醒日期数：{}".format(len(reminders)),
        "",
        "处理明细：",
    ] + (log or ["  (无异常)"])
    with open(os.path.join(args.output_dir, "run_log.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(log_lines))

    print("完成：{}".format(out_xlsx))
    if wrote_ics:
        print("      {}".format(out_ics))
    print("      记录 {} 条，节点 {} 个，提醒 {} 天".format(len(records), len(nodes), len(reminders)))
    return out_xlsx


def run_full(args, log):
    input_path = select_input(args, log)
    log.append("模式：full（全量重新生成）")
    log.append("实际使用输入文件：{}".format(os.path.abspath(input_path)))

    input_wb = load_workbook(input_path, read_only=True, data_only=True)
    generated = SHEET_VERT in input_wb.sheetnames
    input_wb.close()

    df, ncols = load_source(input_path, log)
    persons = load_persons(input_path, log)

    existing_records = read_existing_vertical(input_path) if generated else None
    existing_base = read_existing_base(input_path) if generated else None
    edited_nodes = read_existing_nodes(input_path) if generated else None

    if generated and existing_records:
        log.append("检测到输入为已生成文件，保留『{}』中的人工修改工时。".format(SHEET_VERT))
        records = existing_records
        base = existing_base or {}
        skip_codes = set()
        if not base:
            records, meta, base, skip_codes = build_schedule(df, ncols, persons, log, args.non_interactive)
    else:
        records, meta, base, skip_codes = build_schedule(df, ncols, persons, log, args.non_interactive)

    nodes = build_nodes(df, ncols, persons, log, edited=edited_nodes, skip_codes=skip_codes)
    reminders = build_reminders(nodes)
    emit(records, base, persons, df, nodes, reminders, input_path, args, log, "full")


def run_append(args, log):
    base_path = args.base
    if base_path:
        if not os.path.exists(base_path):
            print("错误：指定的 --base 文件不存在：{}".format(base_path))
            sys.exit(1)
    else:
        base_path = find_latest_base(args.output_dir, log)
        if base_path is None:
            print("错误：未在 {} 找到基线排班文件（WP周排班_*.xlsx，含『{}』）。请先运行 full 模式。".format(
                args.output_dir, SHEET_VERT))
            sys.exit(1)
    log.append("模式：append（增量追加）")
    log.append("基线文件：{}".format(os.path.abspath(base_path)))

    input_path = select_input(args, log)
    log.append("实际使用输入文件：{}".format(os.path.abspath(input_path)))

    df, ncols = load_source(input_path, log)
    persons = load_persons(input_path, log)
    base_persons = load_persons(base_path, [])
    seen, persons_all = set(), []
    for p in persons + base_persons:
        if p["name"] not in seen:
            seen.add(p["name"])
            persons_all.append(p)

    base_records = read_existing_vertical(base_path) or []
    existing_keys = {(r["code"], r["wp"]) for r in base_records}
    log.append("保留的已有项目数（Code+单号）：{}".format(len(existing_keys)))

    # 新项目：按 Engagement Code + WP服务单号 对比基线
    c_code = ncols.get(norm_name("Engagement Code"))
    c_wp = ncols.get(norm_name("WP服务单号"))
    c_eng = ncols.get(norm_name("Engagement Name"))
    new_keys, seen_keys, key_name = [], set(), {}
    if c_code:
        for _, r in df.iterrows():
            code = to_text(r[c_code])
            if code == "":
                continue
            wp = to_text(r[c_wp]) if c_wp else ""
            k = (code, wp)
            if k in seen_keys:
                continue
            seen_keys.add(k)
            key_name[k] = to_text(r[c_eng]) if c_eng else ""
            if k not in existing_keys:
                new_keys.append(k)
    if not new_keys:
        print("append：未发现新增项目（按 Engagement Code + WP服务单号 对比基线），不输出文件。")
        return

    log.append("新增项目（{} 个）：".format(len(new_keys)))
    for k in new_keys:
        log.append("  - {} {} {}".format(k[0], k[1], key_name.get(k, "")))

    reserved = defaultdict(float)
    for r in base_records:
        reserved[(r["staff"], r["week"])] += r["hours"]

    new_records, meta, new_base, skip_codes = build_schedule(
        df, ncols, persons_all, log, args.non_interactive,
        only_keys=set(new_keys), reserved=reserved)

    records = base_records + new_records
    base = merge_base(read_existing_base(base_path) or {}, new_base)

    # 关键节点：既有项目以 WP服务单为准（有值且不同则更新并记录；为空保留基线）；新增项目取 WP服务单
    existing_codes = {k[0] for k in existing_keys}
    base_nodes = read_existing_node_rows(base_path)
    base_by_key = {(n["code"], n["type"]): n for n in base_nodes}
    wp_nodes = build_nodes(df, ncols, persons_all, log, edited=None,
                           skip_codes=set(), only_keys=None)
    wp_by_key = {(n["code"], n["type"]): n for n in wp_nodes}

    merged = {}
    for key, bnode in base_by_key.items():
        wnode = wp_by_key.get(key)
        if wnode is not None and wnode.get("date") is not None:
            if wnode["date"] != bnode.get("date"):
                log.append("  DDL 更新：{} {} {} → {}（以 WP服务单为准）".format(
                    key[0], key[1], bnode.get("date"), wnode["date"]))
            node = dict(wnode)
            if not node.get("name"):
                node["name"] = bnode.get("name", "")
            merged[key] = node
        else:
            merged[key] = bnode
    # 既有项目在 WP服务单 新增的节点（基线没有）→ 以 WP服务单为准补入
    for key, wnode in wp_by_key.items():
        if key[0] in existing_codes and key not in base_by_key and wnode.get("date") is not None:
            log.append("  DDL 新增节点：{} {} {}（以 WP服务单为准）".format(key[0], key[1], wnode["date"]))
            merged[key] = wnode
    # 新增项目（排除数据质量跳过的）
    for key, wnode in wp_by_key.items():
        if key[0] in existing_codes or key[0] in skip_codes:
            continue
        merged.setdefault(key, wnode)

    nodes = sorted(merged.values(), key=lambda x: (x["date"], x["code"], x["type"]))
    reminders = build_reminders(nodes)

    emit(records, base, persons_all, df, nodes, reminders, input_path, args, log,
         "append", base_path=base_path)


def main():
    ap = argparse.ArgumentParser(description="WP 周排班生成器")
    ap.add_argument("--input", default=None, help="指定 WP服务单.xlsx 路径（可选，指定则优先使用）")
    ap.add_argument("--input-dir", default="input", help="扫描输入目录，默认 input")
    ap.add_argument("--output-dir", default="output", help="输出目录")
    ap.add_argument("--mode", choices=["full", "append"], default="full",
                    help="生成模式：full 全量重新生成(默认) / append 增量追加")
    ap.add_argument("--base", default=None,
                    help="append 模式的基线排班 Excel；缺省自动取 output-dir 下最新的 WP周排班_*.xlsx")
    ap.add_argument("--non-interactive", action="store_true", help="非交互模式")
    ap.add_argument("--remind-mode", choices=["outlook", "popup", "both", "ics"], default="both",
                    help="提醒输出：outlook 写Outlook / popup 只生成弹窗脚本 / both 两者(默认) / ics 只生成.ics")
    args = ap.parse_args()

    log = []
    os.makedirs(args.output_dir, exist_ok=True)
    if args.mode == "append":
        run_append(args, log)
    else:
        run_full(args, log)


if __name__ == "__main__":
    main()
