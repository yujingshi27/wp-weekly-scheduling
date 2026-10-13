---
name: wp-weekly-scheduling
description: 从 WP 服务单 Excel 动态读取项目信息和人员清单，按自然周生成周排班宽表与纵表、监控校验 sheet（含公式）和项目关键节点 .ics 日历提醒。当用户提供 WP服务单.xlsx 并要求生成排班、监控 DDL、更新项目进度时使用。
---

# WP 周排班生成器 (wp-weekly-scheduling)

## 何时触发

当用户提供 **WP服务单.xlsx**（文件名不固定），并要求以下任一事项时使用：

- 生成 / 更新 **周排班**（宽表、纵表）
- 生成 **监控与校验 sheet**（周校验、项目工时校验）
- 生成 **项目关键节点** 与 **DDL 日历提醒**（.ics）
- 更新项目进度 / 排班

## 重要：由脚本执行，不要让 LLM 直接生成 Excel

所有计算与写入必须由 `scripts/generate_schedule.py` 完成。LLM 只负责调用脚本、读取
`run_log.txt` 汇报结果、以及在用户需要时解释规则。**不要**手写 Excel 或代替脚本做
工时计算。

## 运行方式（双击即可，无需记参数）

| 一键脚本 | 作用 |
| --- | --- |
| **`run_full.bat`** | **全量重排**：从 `input\` 的服务单重新生成（会覆盖输出 Excel 中的人工修改） |
| **`run_append.bat`** | **增量追加**：保留已有排班，只追加 `input\` 中的新项目 |
| `run.bat` | 全量快捷方式（等价 `run_full.bat`） |

> 修改 DDL 请改 **`WP服务单.xlsx`**，不要改输出 Excel。append 下，已有项目的 DDL 以 WP服务单为准更新（WP服务单为空时保留基线日期）。

等价命令（在 skill 目录下）：

```bat
python scripts\generate_schedule.py --mode full   --input-dir input --output-dir output
python scripts\generate_schedule.py --mode append --input-dir input --output-dir output
```

参数：

| 参数 | 说明 |
| --- | --- |
| `--mode` | `full`（默认，全量重新生成）或 `append`（增量追加到基线排班） |
| `--base` | 可选，append 模式的基线排班 Excel 路径；缺省自动取 `--output-dir` 下最新的 `WP周排班_*.xlsx` |
| `--input` | 可选；指定则优先使用该文件，文件不存在会直接报错 |
| `--input-dir` | 扫描目录，默认 `input`。未指定 `--input` 时扫描该目录下所有 `.xlsx`，自动识别含 `WP服务单` sheet 且表头含 `Engagement Code` / `剩余hours` 的文件 |
| `--output-dir` | 输出目录，默认 `output` |
| `--remind-mode` | `outlook`（写 Outlook）/ `popup`（只生成本地弹窗脚本）/ `both`（两者，默认）/ `ics`（只生成 `.ics`） |
| `--non-interactive` | 非交互模式；缺少预审开始时间时默认跳过，不询问 |

### 日历提醒（Outlook / 本地弹窗 / .ics）

- **Outlook（`--remind-mode outlook|both`）**：用 `pywin32` 调本机 Outlook 日历，先删所有带 `WP-DDL-` 标识的事件，再按「日历提醒汇总」逐日写入（标题=汇总标题、10:00–10:30、正文=汇总详情、提前 0 分钟提醒）。
  若 Outlook/pywin32 不可用 → 记入 `run_log.txt` 并**自动降级生成 `.ics`**。
- **本地弹窗脚本（`--remind-mode popup|both`）**：生成到 `scripts\`：
  - `daily_check.py` / `daily_check.bat`：检查未来 7 天内的提醒日期，有则弹窗「WP DDL 提醒」（`plyer`→`win10toast`→PowerShell 兜底，全本地）。
  - `register_task.bat`：一键注册 Windows 任务计划 `WP_DDL_DailyCheck`（**每天 10:30** 运行 `daily_check.bat`）。
  - `unregister_task.bat`：一键删除该任务。
- **`.ics`（`--remind-mode ics`）**：只生成 `项目关键节点.ics`。

输入文件**不固定文件名**，动态扫描选择规则：

- 0 个候选：报错并列出扫描目录下所有 `.xlsx` 文件名，提示检查格式；
- 1 个候选：直接使用并记录文件名；
- 多个候选：交互模式让用户输入序号选择，非交互模式（`--non-interactive`）取修改时间最新的并写日志。

### 增量追加模式（append）

用于在**已人工修改过的排班表**上新增项目，保留已有项目的人工调整：

```bat
run_append.bat
```

或对灵码说“用 append 模式跑 WP 排班”。逻辑：

1. 读取基线（`--base` 或 `output` 下最新 `WP周排班_*.xlsx`）的 `周排班-纵表`、`项目基础`、`项目关键节点`、隐藏 `人员清单`。
2. 按 **Engagement Code + WP服务单号** 对比，找出不在基线纵表中的**新增项目**；无新增项目则**不输出任何文件**并退出。
3. 只对新增项目排班，插入纵表；**已有项目行原样保留、不重算、不覆盖**；新项目只使用每周**剩余容量**（占用基线已用工时后），冲突写入 `run_log.txt`，不改动已有行。
4. 关键节点 DDL：既有项目按优先级取日期——**WP服务单有值且与基线不同 → 以 WP服务单为准更新并记日志**；**WP服务单为空 → 保留基线（含人工修改）**；新增项目取 WP服务单。改完 WP服务单跑 append 即可刷新 Outlook/`.ics`/`daily_check` 使用的日期。
5. 宽表 / 周校验 / 项目工时校验 / 日历提醒汇总全部重建，公式覆盖含新项目的所有行；按 `--remind-mode` 重写 Outlook/弹窗/`.ics`。
6. 输出新的 `WP周排班_新时间戳.xlsx`，**不覆盖基线**；`run_log.txt` 记录新增项目列表、DDL 更新、容量冲突、保留的已有项目数。

依赖：`pip install -r requirements.txt`（pandas、openpyxl；Windows 下加 `pywin32` 供 Outlook 写入，可选）。

## 输入

- **WP服务单 sheet**：按**表头名称动态定位**，列顺序可变化、无关列可增减，**不得按列号硬编码**。
  需抓取：`Client Name, Engagement Code, Engagement Name, 订单编号, WP服务单号, Audit EIC,
  Audit Report Date, Outlook Hours, 剩余hours, 项目类型, 预审开始时间, 预审交付时间, 年审开始时间, 年审交付时间`。
  （**不再使用预计桶**如 `预审（10-12月）预计hours / 1-3月预计hours / 4-6月预计hours`；若存在则忽略。）
  （表头容错：忽略空格，`WP 服务单号` 与 `WP服务单号` 均可。）
- **参与人员**：按表头前缀 `参与人员` 识别，有多少识别多少（5 人、20 人皆可）。
- **人员清单 sheet**（可选）：表头 `Name, Grade, hours占比, 优先项目类型`。读取失败或缺失时回退到内置默认清单
  （Marsh/manager、Vivi/Senior、Yuan/Luna/Rainbow/staff、Intern/intern）。`Grade` 支持 `manager/senior/staff/intern`（不区分大小写）。
  `优先项目类型` 为可选列（如 `公募`、`PE`，空=无偏好）。
  `Staff Name` 输出直接使用清单 `Name` 原值，不做映射。
  短名匹配（如 `Vivi`、`yuan`）与清单 `Name` 大小写不敏感地按首段前缀匹配。

## 规则

1. **自然周**：EY 财年 **周六 → 周五**。周期标识 `WeekNN (YYYY.MM.DD)`，日期为该周周六。
   WeekNN 由内置财年算法计算（FY 起始 = 距 (FY-1).7.1 最近的周六），无需外部文件。
2. **工时基准**：**周标准工时** 忙季(1/2/3月) **60h**、淡季 **40h**；**硬上限** = 标准 × 1.2（**72h / 48h**）。
3. **角色配额与目标**：manager 固定 **8%** 均摊（不参与优化、不受上限约束）；senior **22%**；staff+intern 共享 **70%** 池。
   角色周目标：staff = 标准（60/40）、senior = 0.85×标准（51/34）、intern = **固定 40h/周**。缺口并入池。
   未在人员清单出现的人不分配工时。
   **优先项目类型（软偏好）**：若人员清单填了 `优先项目类型`，同一员工同一周有多个项目时，**其偏好类型的项目优先于交付日期**
   （仅作用于 70% 池；manager/senior 不受影响；空=无偏好，单值精确匹配）。
4. **工时预算与阶段分配**：**不再使用预计桶**。每行项目**预算 = 剩余hours**；**预审比例**（占预算）默认 **30%**（按 `项目类型` 可配，见代码 `PRE_RATIO_BY_TYPE`/`DEFAULT_PRE_RATIO`，后续人工调整）：
   - 有预审（`预审开始时间` 存在）：`预审 = 比例×剩余hours`，`年审 = 剩余hours − 预审`；
   - 无预审：`年审 = 100%×剩余hours`，不排预审。
   工时按月份天数比例拆分，再在月内按自然周天数均分；首/末周按落在月内的天数折算；同周开始交付则全落该周。
   两阶段之和 = `剩余hours`，保证项目 diff<5h（改纵表工时后人工微调）。
5. **周容量填充与项目完成优先级**：按周、按“项目完成优先级”（阶段交付日期升序，同日期按 `Engagement Code` 升序，再按输入顺序）
   为每个项目填充工时。**优先按目标工时排**；装不下则**顺延到后续周**；仅在该项目**末周**才允许**上浮到硬上限**兜底，绝不超硬上限。
   被挤占的工时无限顺延（可跨月），保证项目 `排班合计` 与 `剩余hours` 的 diff < 5h。
   `run_log.txt` 记录项目优先级排序、上浮触发（人/周）、因优先级低“剩余容量不足”、周超硬上限警告。
6. **intern 规则**：仅在 **1–3 月**参与排班；与 staff 共享 70% 池——同一项目同一周先把 staff 排到目标，剩余才分给 intern；
   intern **每周上限 40h**（目标与硬上限均为 40h）。staff 吃满池导致 intern 无活可排时记入 `run_log.txt`。
7. **报告日约束**：读取 `Audit Report Date`（报告日），任一阶段**交付日 > 报告日** 视为疑似日期有误，交互批量询问是否跳过
   （`--non-interactive` 默认跳过）。跳过 = **整项目**不排班并从关键节点/日历剔除；不跳过则排班末端**截断到报告日所在周**并记日志。
8. **缺失/异常日期**：`NA`/空视为缺失；模糊日期（如 `2026年11月初`）近似为当月 5/15/25 日并记录日志；
   交付早于开始 → 跳过该阶段并记录；交付缺失 → 整段落在开始周，超上限部分向后滚排。
   预审与年审交叉校验：**任一预审时间晚于任一年审时间**（`max(预审) > min(年审)`）视为数据有误，交互批量询问，`--non-interactive` 默认跳过。
9. 使用 `--input` 指向**已生成的工作簿**时，脚本会读取其中的 `源数据-WP服务单`、`周排班-纵表`
   （保留人工改过的工时）与 `项目关键节点`（保留人工改过的日期），刷新公式与 .ics，支持“改完重跑”。

## 输出（`WP周排班_YYYYMMDD_HHMM.xlsx`，不覆盖原文件）

| Sheet | 内容 |
| --- | --- |
| 周排班-纵表 | **唯一数据源**，人工只改这里。列：项目类型/Client/Code/Engagement Name/订单编号/WP服务单号/Staff Name/周期/hours |
| 周排班-宽表 | A–G 固定表头，H 起每自然周一列，单元格为 `SUMIFS` 公式引用纵表；可直接粘贴到排班助手 |
| 周校验 | 人员×周期：该周合计(SUMIFS)、是否忙季、周标准工时、UT比例(`=ROUND(合计/40,4)`，**百分比**，以 40h 为基准)、是否超硬上限(标红)。intern 标准 40、超硬>40；非 intern 标准 淡季 40/忙季 60、超硬阈值 淡季 48 / 忙季 78。公式保留 |
| 项目工时校验 | 按项目，剩余hours / 排班合计（SUMIFS）/ diff / diff<5h（标黄），公式保留 |
| 项目关键节点 | 4 类节点 + 4 个提醒列公式（-7/-5/-3/当天），按日期升序 |
| 日历提醒汇总 | 按提醒日期聚合，标题 `DDL汇总：YYYY-MM-DD（N项）`，详情换行（唯一的 DDL 提醒方式） |

（已删除原「本周待办与DDL预警」sheet；DDL 提醒仅通过 `项目关键节点.ics` 提供。）

**格式**：纵表按 **周期 → 订单编号** 升序；周校验按 **周期 → Staff Name** 升序。字体**统一黑体**（表头黑体粗体白字）；表头行高按内容自适应；宽表 H 起周期列固定宽度 11，所有 sheet 统一套用边框与列宽美化。

另生成：

- `项目关键节点.ics`（或按 `--remind-mode` 写入 Outlook）：每个提醒日期一个事件（10:00–10:30）；
  Outlook 事件带 `WP-DDL-{日期}-{阶段}` 标识，重跑先删后写、不重复不残留。
- 本地弹窗脚本：`scripts\daily_check.py`、`daily_check.bat`、`register_task.bat`、`unregister_task.bat`。
- `run_log.txt`：处理信息、模糊日期近似、跳过项目、项目优先级排序、上浮触发（人/周）、intern 未排满、报告日超期/截断、**DDL 更新（append）**、Outlook 删除/写入数或**降级 .ics**、数据质量异常、周超硬上限与工时偏差警告。
- 隐藏辅助 sheet：`源数据-WP服务单`、`项目基础`、`人员清单`（用于公式引用与重跑回读）。
