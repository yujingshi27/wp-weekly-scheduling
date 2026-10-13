# WP 周排班生成器（wp-weekly-scheduling）

从 **WP 服务单.xlsx** 动态读取项目信息与人员清单，按自然周（EY 财年 周六→周五）生成：

- **周排班-纵表**（唯一数据源，人工只改这里）＋ **周排班-宽表**（SUMIFS 公式引用纵表）
- **周校验**（标准工时 / UT 比例 / 超硬上限）＋ **项目工时校验**（diff<5h）
- **项目关键节点**（4 类节点 + 提醒列公式）＋ **日历提醒汇总**
- **Outlook 日历事件 / 本地弹窗脚本 / `.ics`** 三种 DDL 提醒
- **增量追加（append）**：在人工修改过的排班表上只追加新项目

---

## 目录结构

```
.opencode/skills/wp-weekly-scheduling/
├─ SKILL.md               # opencode 技能说明（触发条件、规则、调用方式）
├─ README.md              # 本文件
├─ requirements.txt       # pandas、openpyxl、（Windows 可选）pywin32
├─ run_full.bat           # 双击：全量重排
├─ run_append.bat         # 双击：增量追加
├─ run.bat                # 全量快捷方式（调用 run_full.bat）
├─ input/                 # 放 WP服务单.xlsx（文件名不限）
├─ output/                # 生成的 Excel / .ics / run_log.txt
└─ scripts/
   ├─ generate_schedule.py # 主脚本（所有计算与写 Excel）
   ├─ daily_check.py       # 每日弹窗检查（运行时自动生成）
   ├─ daily_check.bat      # 由任务计划调用（运行时自动生成）
   ├─ register_task.bat    # 一键注册每日 10:30 任务
   └─ unregister_task.bat  # 一键删除该任务
```

> `daily_check*` 与 `register/unregister_task.bat` 会在每次运行时自动生成到 `scripts/`。

---

## 安装

```bat
pip install -r requirements.txt
```

依赖：`pandas`、`openpyxl`；Windows 上安装 `pywin32` 才能写 Outlook 日历（缺失时自动降级为 `.ics`）。

---

## 快速开始（双击即可，无需记参数）

| 脚本 | 作用 |
| --- | --- |
| **run_full.bat** | 全量重排：从 `input\` 重新生成（**会覆盖输出 Excel 中的人工修改**） |
| **run_append.bat** | 增量追加：保留已有排班，只追加 `input\` 中的新项目 |
| run.bat | 全量快捷方式（等价 `run_full.bat`） |

1. 把 WP 服务单放到 `input\`（文件名不限，需含 `WP服务单` sheet）。
2. 双击对应 bat。
3. 结果在 `output\`：`WP周排班_YYYYMMDD_HHMM.xlsx`、`项目关键节点.ics`、`run_log.txt`。

> **改 DDL 请改 `WP服务单.xlsx`，不要改输出 Excel。** append 下既有项目 DDL 以 WP服务单为准更新，WP服务单为空则保留基线日期。

---

## 命令行参数

```bat
python scripts\generate_schedule.py --mode full   --input-dir input --output-dir output
python scripts\generate_schedule.py --mode append --input-dir input --output-dir output
```

| 参数 | 说明 |
| --- | --- |
| `--mode` | `full`（默认，全量重生成）或 `append`（增量追加） |
| `--base` | append 的基线排班 Excel；缺省自动取 `--output-dir` 下最新 `WP周排班_*.xlsx` |
| `--input` | 可选，指定则优先使用该文件（不存在直接报错） |
| `--input-dir` | 扫描目录，默认 `input`；自动识别含 `WP服务单` sheet 且表头含 `Engagement Code`/`剩余hours` 的文件 |
| `--output-dir` | 输出目录，默认 `output` |
| `--remind-mode` | `outlook` / `popup` / `both`（默认）/ `ics` |
| `--non-interactive` | 非交互模式 |

---

## 输入

- **WP服务单 sheet**：按表头名称**动态定位**（列顺序可变、无关列可增删）。抓取：`Client Name, Engagement Code, Engagement Name, 订单编号, WP服务单号, Audit EIC, Audit Report Date, Outlook Hours, 剩余hours, 项目类型, 预审开始时间, 预审交付时间, 年审开始时间, 年审交付时间`。
  - **不再使用预计桶**（`预审/1-3/4-6预计hours`）；若存在则忽略。
- **参与人员**：按表头前缀 `参与人员` 识别，任意数量（5、20 皆可）。
- **人员清单 sheet**（可选）：`Name, Grade, hours占比, 优先项目类型`。缺失时回退内置默认清单。
  - `Grade`：manager / senior / staff / intern（不区分大小写）。
  - `优先项目类型`：软偏好（如 `公募`、`PE`），空=无偏好。
  - `Staff Name` 输出直接用清单 `Name` 原值；短名（如 `Vivi`、`yuan`）与清单 `Name` 大小写不敏感地按首段前缀匹配。

---

## 主要规则

1. **自然周**：EY 财年 周六→周五；标识 `WeekNN (YYYY.MM.DD)`（日期为该周周六），内置财年算法，无需外部文件。
2. **工时基准**：周标准工时 忙季(1/2/3月) 60h / 淡季 40h；硬上限 = 标准×1.2（72/48）。
3. **角色**：manager 8%（均摊、不受上限）；senior 22%；staff+intern 共享 70% 池。
   - 周目标：staff=标准、senior=0.85×标准、intern=固定 40h/周。
   - **优先项目类型**：同一员工同周多项目时，偏好类型优先于交付日期（仅作用于 70% 池）。
4. **工时预算**：每行项目 **预算 = 剩余hours**；**预审比例默认 30%**（按项目类型可配 `PRE_RATIO_BY_TYPE`），`预审=比例×预算`、`年审=预算−预审`；**无预审**则年审=100%。两阶段合计=剩余hours → diff<5h。比例可用输出纵表工时人工微调。
5. **周容量填充**：按项目完成优先级（阶段交付日期升序→Code→输入顺序）填工时；优先按目标、装不下顺延到后续周、仅末周上浮到硬上限（绝不超）。
6. **intern**：仅 1–3 月参与；每周上限 40h；staff 先到目标，剩余才给 intern。
7. **报告日约束**：`Audit Report Date`；任一阶段交付日>报告日视为疑似有误，交互询问是否跳过（非交互默认跳过），不跳过则截断到报告日所在周。
8. **数据质量**：预审时间晚于年审时间 → 疑似有误，交互询问是否跳过（非交互默认跳过）。
9. **append**：按 `Engagement Code + WP服务单号` 找新增项目；已有行原样保留、不重算；关键节点既有项目**WP服务单有值且不同则以 WP服务单为准更新**、为空保留基线。

---

## 输出

`WP周排班_YYYYMMDD_HHMM.xlsx`（不覆盖原文件）：

| Sheet | 内容 |
| --- | --- |
| 周排班-纵表 | **唯一数据源**，人工只改这里。列：项目类型/Client/Code/Engagement Name/订单编号/WP服务单号/Staff Name/周期/hours |
| 周排班-宽表 | A–G 固定表头，H 起每自然周一列，SUMIFS 公式引用纵表；可粘贴到 SG 排班助手 |
| 周校验 | 人员×周期：该周合计、是否忙季、周标准工时、UT比例（百分比，以 40h 为基准）、是否超硬上限（标红） |
| 项目工时校验 | 按项目：剩余hours / 排班合计 / diff / diff<5h（标黄） |
| 项目关键节点 | 4 类节点 + 4 个提醒列公式（-7/-5/-3/当天） |
| 日历提醒汇总 | 按提醒日期聚合，标题 `DDL汇总：YYYY-MM-DD（N项）` |

- 纵向排序：**周期 → 订单编号**；周校验排序：**周期 → Staff Name**。
- 字体：**统一黑体**；表头行高按内容自适应。
- 隐藏辅助 sheet：`源数据-WP服务单`、`项目基础`、`人员清单`（用于公式引用与重跑回读）。

### 日历提醒（`--remind-mode`）

- **outlook / both**：用 `pywin32` 写本机 Outlook 日历；重跑**先删**所有带 `WP-DDL-` 标识的事件再重写；不可用时**自动降级 `.ics`**。
- **popup / both**：生成 `scripts\daily_check.py` + `daily_check.bat` + `register_task.bat` + `unregister_task.bat`。
  - `daily_check.py`：检查未来 7 天提醒，有则弹窗「WP DDL 提醒」（`plyer`→`win10toast`→PowerShell 兜底）。
  - 双击 `register_task.bat` 注册任务 `WP_DDL_DailyCheck`（**每天 10:30**）；`unregister_task.bat` 删除。
- **ics**：只生成 `项目关键节点.ics`。

`run_log.txt` 记录：跳过项目、优先级排序、上浮、intern 未排满、报告日截断、DDL 更新、Outlook/降级、数据质量异常、周超硬上限与工时偏差等。

---

## 常见问题

- **找不到输入文件**：确认 `input\` 下有含 `WP服务单` sheet 且表头含 `Engagement Code` / `剩余hours` 的 `.xlsx`。
- **Outlook 没写进去**：多为未安装 `pywin32` 或本机无 Outlook；日志会记录并自动用 `.ics` 兜底。
- **排班与剩余hours 不一致**：项目某阶段日期缺失导致该阶段被跳过，见 `run_log.txt` 与「项目工时校验」标黄项。
- **改了排班想保留**：用 `run_append.bat`（只追加新项目）；需要彻底重排才用 `run_full.bat`。

---
---

# WP 本地工作台（workbench）

一个**只在本机运行**的网页工作台，把「待办管理」和「工时筛选」合在一起，界面为浅色卡片风。
数据存本机 SQLite，不需要联网、不需要密码、不需要外部 API/AI 网站。
本机运行说明见 `workbench\使用说明.txt`。

> 目录：`C:\AIProjects\3-wp-weekly-scheduling\workbench\`

## 功能概览

- **待办管理**：状态分区（今日得做 / 待完成 / 卡着 / 等反馈 / 归档）、标签、备注、分类（日常任务 / 临时任务 / 常规重复 / DDL）、**四象限**（重要且紧急 / 重要不紧急 / 紧急不重要 / 不紧急不重要）、截止日期；新建时状态直接生效，新建/编辑/开机都会按截止日期流转状态；常规重复事项完成后自动生成下一期并归档；归档满 7 天自动删除、支持按名称模糊搜索；列表按「截止日期 + 四象限」排序。
- **工时筛选**：按项目（多 code + 多 manager 模糊匹配）或按人；汇总表 + 项目合计区（合计工时 vs Outlook Hours 预算，超预算标红，按 code 升序）；可导出 Excel。
- **DDL 自动进待办**：从排班输出的「项目关键节点」sheet 读取，一条 DDL = 一条待办；DDL 日期改了会更新，**已完成的不动**。
- **数据源配置**：首次手动指定三个文件（Timesheet / WP服务单 / 排班输出），支持分别「替换」；文件名不固定，全部按表头名动态识别。
- **自动备份**：每周一次，复制 `todo.db` 到 `backup\`，保留最近 7 份。

## 目录结构

```
workbench\
├─ 启动工作台.bat        # 双击启动（调 .venv 的 Python）
├─ WP工作台.exe          # 打包后的独立程序（自带运行环境）
├─ build_exe.bat         # 重新打包 exe
├─ 使用说明.txt          # 面向非技术用户的说明
├─ requirements.txt      # flask、openpyxl、lxml
├─ config.json           # 数据源路径、Timesheet 列映射、上次备份周、端口
├─ data\todo.db          # SQLite：待办 + 标签 + Timesheet 缓存 + 预算缓存
├─ backup\               # 每周备份 todo_YYYYMMDD.db（保留 7 份）
├─ exports\              # 导出的 Excel
└─ app\
   ├─ app.py             # Flask 入口（启动后自动开浏览器）
   ├─ config.py          # 路径 / 配置读写（兼容 PyInstaller 冻结）
   ├─ db.py              # SQLite 建库与连接
   ├─ todo_service.py    # 待办 CRUD、状态规则、重复规则、归档清理
   ├─ ddl_service.py     # 从「项目关键节点」同步 DDL 待办
   ├─ timesheet_service.py # 筛选、汇总、项目合计
   ├─ datasource.py      # 表头动态定位、Timesheet 流式导入、预算导入
   ├─ backup.py          # 每周备份 + 保留 7 份
   ├─ export.py          # 导出 Excel
   ├─ templates\         # base / index / timesheet / settings
   └─ static\            # style.css、app.js、todo.js、timesheet.js、settings.js
```

## 启动

| 方式 | 说明 |
| --- | --- |
| 双击 `启动工作台.bat` | 用 `workbench\.venv` 的 Python 启动 Flask，并**自动打开浏览器** `http://127.0.0.1:5000`；关闭黑窗口即退出 |
| 双击 `WP工作台.exe` | 打包后的独立程序，行为同上，无需装 Python |
| 重新打包 | 双击 `build_exe.bat`（需联网装 PyInstaller），输出 `WP工作台.exe` |

端口默认 `5000`，被占用时自动顺延。数据源等设置存在 `config.json`。

## 首次设置（三步）

1. **数据源设置** → 填三个文件路径（可「浏览」选，或粘贴路径）→ 保存：
   - Timesheet 文件（工时筛选）
   - WP服务单文件（Outlook Hours 预算）
   - 排班输出文件（DDL / 项目关键节点）
2. **Timesheet 列映射** → 点「读取表头」→ 7 个字段选列 → 保存。默认推荐：

   | 用途 | 推荐列 |
   | --- | --- |
   | 项目编号 | `Engagement Code` |
   | 员工 | `Employee Name` |
   | manager | `COE Manager` |
   | 工时 | `Charge Hours` |
   | 期间 | `Transaction Cycle Date`（可改 `Fiscal Week`） |
   | 描述 | `Charge Description` |
   | 项目名称 | `Engagement Name` |

3. **导入数据** → 「刷新 Timesheet」（48 万行首刷约 3 分钟，之后筛选秒出）、「刷新预算」（几秒）。

## 工时筛选输出

- 汇总表字段：项目名称 / code / staffname / manager name / hours / 期间 / 描述。
- 项目合计区：按 code 升序，显示 合计工时 vs 预算（Outlook Hours），超预算「是否超预算」标红。
- 导出文件名：按项目 `工时筛选_YYYYMMDD_首个code等.xlsx`；按人 `工时筛选_YYYYMMDD_首个manager等.xlsx`（多个只写第一个加「等」）。
- 筛选结果为空 → 页面红字提示，重新输入，不退出。

## 待办规则要点

- 新建时选择的状态**直接生效**（选“卡着”就进卡着区）；未选默认「待完成」。
- 状态流转：**新建 / 编辑 / 打开工作台**时都会检查——截止日 ≤ 今天且状态为 待完成/卡着/等反馈 → 「今日得做」；已过期未完成 → 保持「今日得做」并标红置顶。
- 分类：日常任务 / 临时任务（普通）、常规重复、DDL；旧数据的「其他」在首次启动时自动迁移为「日常任务」。
- 四象限：未设置 / 重要且紧急 / 重要不紧急 / 紧急不重要 / 不紧急不重要。普通新建默认未设置；DDL 自动生成默认「重要且紧急」；常规重复新一期继承上一期。
- 排序：`due_date` 升序（无截止日排最后）→ 同日期按象限优先级 → id；逾期因日期最早自然置顶。
- 常规重复：每天 / 每周几 / 每月几号 / 工作日；完成后归档原条并自动生成下一期；把分类从「常规重复」改成别的即停止重复。
- DDL 待办按 **项目 code + 节点类型** 识别；日期改动会更新，已完成的不改。备注格式：`项目：XX（Code47129402）| 节点：预审开始`。
- **数据库迁移**：`init_db()` 会对旧库自动 `ALTER TABLE ADD COLUMN quadrant`，无需手动处理。

## 数据与备份

- 待办数据在 `data\todo.db`；每次打开检查「本周是否已备份」，未备则复制到 `backup\todo_YYYYMMDD.db`，保留最近 7 份。
- 归档数据保留 7 天后彻底删除（从进归档那天算）。

## 技术要点

- **框架**：Flask（服务端渲染 Jinja + 少量原生 JS），全部资源本地化，无 CDN。
- **大数据**：Timesheet 的 sheet XML 解压后约 **2.5 GB**（48 万行 × 135 列），openpyxl/expat 会因 2GB 上限 OOM，故改用 **lxml 流式解析**（逐行 `clear`），只抽取映射的列，内存可控，全量导入约 3 分钟。
- **筛选**：条件尽量下推到 SQL（code `IN`、manager `LIKE` 模糊匹配，忽略大小写），展示最多 2000 行，完整数据走导出。
- **打包**：PyInstaller onefile，`--add-data` 打包 `templates`/`static`；`config.py` 用 `resource_path()` 兼容冻结运行，`data`/`backup`/`exports`/`config.json` 生成在 exe 同目录。
- **单实例**：启动先探测 `/api/health`，若工作台已在运行则直接复用并打开浏览器，不重复起服务；导入作业加进程内锁（`JOB_LOCK`）；SQLite 启用 `WAL` + 30s `busy_timeout`，避免并发写锁导致 500。
- **文件选择**：数据源「浏览」调用 PowerShell `System.Windows.Forms.OpenFileDialog`（开发与 EXE 通用，不再多开实例）。
- **排错**：未捕获异常写入 `workbench\logs\error.log` 并在页面显示真实错误信息。
- **启动脚本**：`启动工作台.bat`/`build_exe.bat` 使用 CRLF 换行（LF-only 的 bat 会导致双击无反应）。

## 依赖

```bat
pip install -r workbench\requirements.txt
```

`flask`、`openpyxl`、`lxml`（打包另需 `pyinstaller`）。
