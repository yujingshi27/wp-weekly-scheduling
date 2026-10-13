(function () {
  "use strict";
  const { esc, api } = window.WB;

  const saveSourcesBtn = document.getElementById("save-sources");
  if (!saveSourcesBtn) return;

  const SRC = {
    timesheet: "ds-timesheet",
    wp_service: "ds-wp_service",
    schedule_output: "ds-schedule_output",
  };
  const MAPPING_KEYS = ["project_code", "staff", "manager", "hours", "period", "description", "project_name"];

  function banner(msg, kind) {
    document.querySelectorAll(".page-alert").forEach((n) => n.remove());
    const box = document.createElement("div");
    box.className = "page-alert " + (kind === "error" ? "alert-error" : "alert-ok");
    box.textContent = msg;
    const main = document.querySelector(".container");
    main.insertBefore(box, main.firstChild);
    setTimeout(() => box.remove(), 4500);
  }

  function fillMappingSelects(headers, current) {
    MAPPING_KEYS.forEach((key) => {
      const sel = document.querySelector('.mapping-grid select[data-key="' + key + '"]');
      const saved = (current && current[key]) || "";
      sel.innerHTML = '<option value=""></option>' +
        headers.map((h) => '<option value="' + esc(h) + '">' + esc(h) + "</option>").join("");
      if (saved && headers.includes(saved)) sel.value = saved;
    });
  }

  async function loadSettings() {
    const data = await api("/api/settings");
    Object.keys(SRC).forEach((k) => {
      document.getElementById(SRC[k]).value = (data.datasources && data.datasources[k]) || "";
    });
    fillMappingSelects([], data.timesheet_mapping || {});
    // seed saved values as options so they show before reading headers
    MAPPING_KEYS.forEach((key) => {
      const sel = document.querySelector('.mapping-grid select[data-key="' + key + '"]');
      const saved = data.timesheet_mapping && data.timesheet_mapping[key];
      if (saved) sel.innerHTML = '<option value=""></option><option value="' + esc(saved) + '">' + esc(saved) + "</option>";
      sel.value = saved || "";
    });
    document.getElementById("ts-sheet").innerHTML =
      '<option value="">（默认第一个）</option>' +
      (data.timesheet_sheet ? '<option value="' + esc(data.timesheet_sheet) + '">' + esc(data.timesheet_sheet) + "</option>" : "");
    if (data.timesheet_sheet) document.getElementById("ts-sheet").value = data.timesheet_sheet;
  }

  async function loadStatus() {
    try {
      const s = await api("/api/timesheet/status");
      Object.keys(SRC).forEach((k) => {
        const flag = document.querySelector('.ds-flag[data-flag="' + k + '"]');
        if (!flag) return;
        const exists = s.datasource_exists && s.datasource_exists[k];
        const path = s.datasources && s.datasources[k];
        flag.textContent = path ? (exists ? "文件存在" : "文件不存在") : "未设置";
        flag.className = "ds-flag " + (exists ? "ok" : (path ? "bad" : ""));
      });
      const info = document.getElementById("cache-info");
      info.innerHTML =
        "Timesheet 缓存：" + s.ts_rows + " 条" +
        (s.ts_last_import ? "（更新于 " + esc(s.ts_last_import) + "）" : "") +
        "<br>预算缓存：" + s.budgets + " 个项目" +
        (s.budget_last_import ? "（更新于 " + esc(s.budget_last_import) + "）" : "");
    } catch (e) { /* ignore */ }
  }

  document.querySelectorAll("[data-pick]").forEach((btn) => {
    btn.addEventListener("click", async () => {
      try {
        const data = await api("/api/pick_file", "POST", {});
        if (data.path) {
          document.getElementById(btn.dataset.pick).value = data.path;
        }
      } catch (err) { banner(err.message, "error"); }
    });
  });

  saveSourcesBtn.addEventListener("click", async () => {
    const ds = {};
    Object.keys(SRC).forEach((k) => { ds[k] = document.getElementById(SRC[k]).value; });
    try {
      await api("/api/settings", "POST", { datasources: ds });
      await loadStatus();
      banner("数据源已保存", "ok");
    } catch (err) { banner(err.message, "error"); }
  });

  document.getElementById("load-headers").addEventListener("click", async () => {
    const path = document.getElementById("ds-timesheet").value;
    if (!path) { banner("请先填写 Timesheet 路径", "error"); return; }
    try {
      const sh = await api("/api/settings/sheets", "POST", { source: "timesheet", path: path });
      const sheetSel = document.getElementById("ts-sheet");
      sheetSel.innerHTML = '<option value="">（默认第一个）</option>' +
        sh.sheets.map((s) => '<option value="' + esc(s) + '">' + esc(s) + "</option>").join("");
      const sheet = sheetSel.value || "";
      const hd = await api("/api/settings/headers", "POST", { source: "timesheet", path: path, sheet: sheet });
      fillMappingSelects(hd.headers, {});
      banner("已读取表头，共 " + hd.headers.length + " 列", "ok");
    } catch (err) { banner(err.message, "error"); }
  });

  document.getElementById("ts-sheet").addEventListener("change", async () => {
    const path = document.getElementById("ds-timesheet").value;
    if (!path) return;
    try {
      const hd = await api("/api/settings/headers", "POST", {
        source: "timesheet", path: path, sheet: document.getElementById("ts-sheet").value,
      });
      fillMappingSelects(hd.headers, {});
    } catch (err) { banner(err.message, "error"); }
  });

  document.getElementById("save-mapping").addEventListener("click", async () => {
    const mapping = {};
    MAPPING_KEYS.forEach((key) => {
      mapping[key] = document.querySelector('.mapping-grid select[data-key="' + key + '"]').value;
    });
    try {
      await api("/api/settings", "POST", {
        timesheet_mapping: mapping,
        timesheet_sheet: document.getElementById("ts-sheet").value,
      });
      banner("列映射已保存", "ok");
    } catch (err) { banner(err.message, "error"); }
  });

  document.getElementById("refresh-ts").addEventListener("click", async () => {
    const btn = document.getElementById("refresh-ts");
    btn.disabled = true;
    banner("正在读取 Timesheet，文件较大请耐心等待（约 3 分钟）...", "ok");
    try {
      const data = await api("/api/timesheet/refresh", "POST");
      await loadStatus();
      banner("Timesheet 已刷新：" + data.ts_rows + " 条记录（用时 " + data.seconds + " 秒）", "ok");
    } catch (err) {
      banner(err.message, "error");
    } finally {
      btn.disabled = false;
    }
  });

  document.getElementById("refresh-budget").addEventListener("click", async () => {
    const btn = document.getElementById("refresh-budget");
    btn.disabled = true;
    try {
      const data = await api("/api/timesheet/refresh_budget", "POST");
      await loadStatus();
      banner("预算已刷新：" + data.budgets + " 个项目", "ok");
    } catch (err) {
      banner(err.message, "error");
    } finally {
      btn.disabled = false;
    }
  });

  loadSettings().then(loadStatus).catch((e) => banner(e.message, "error"));
})();
