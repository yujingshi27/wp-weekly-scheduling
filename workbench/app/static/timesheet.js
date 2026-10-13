(function () {
  "use strict";
  const { esc, api } = window.WB;

  const root = document.getElementById("btn-filter");
  if (!root) return;

  const codeInputs = document.getElementById("code-inputs");
  const managerInputs = document.getElementById("manager-inputs");
  const codeGroup = document.getElementById("code-group");
  const statusLine = document.getElementById("status-line");
  const resultSection = document.getElementById("result-section");
  const resultTable = document.getElementById("result-table");
  const resultCount = document.getElementById("result-count");
  const totalsSection = document.getElementById("totals-section");
  const totalsTable = document.getElementById("totals-table");

  let lastQuery = null;

  function addInput(container, placeholder) {
    const row = document.createElement("div");
    row.className = "multi-row";
    row.innerHTML =
      '<input type="text" placeholder="' + placeholder + '">' +
      '<button type="button" class="btn danger btn-sm">删除</button>';
    row.querySelector("button").addEventListener("click", () => {
      if (container.children.length > 1) row.remove();
      else row.querySelector("input").value = "";
    });
    container.appendChild(row);
    return row;
  }

  function ensureOne(container, placeholder) {
    if (!container.children.length) addInput(container, placeholder);
  }

  ensureOne(codeInputs, "如 47129402");
  ensureOne(managerInputs, "如 A 或 B");

  document.getElementById("add-code").addEventListener("click", () => addInput(codeInputs, "如 47129402"));
  document.getElementById("add-manager").addEventListener("click", () => addInput(managerInputs, "如 A 或 B"));

  function valuesOf(container) {
    return Array.from(container.querySelectorAll("input"))
      .map((i) => i.value.trim())
      .filter(Boolean);
  }

  function currentMode() {
    return document.querySelector('input[name=mode]:checked').value;
  }

  function syncMode() {
    codeGroup.classList.toggle("hidden", currentMode() === "person");
  }
  document.querySelectorAll('input[name=mode]').forEach((r) => r.addEventListener("change", syncMode));
  syncMode();

  function fmt(n) {
    if (n === null || n === undefined || n === "") return "—";
    return Number(n).toLocaleString("zh-CN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function clearResults() {
    resultSection.classList.add("hidden");
    totalsSection.classList.add("hidden");
    document.querySelectorAll(".alert-error, .alert-ok").forEach((n) => n.remove());
  }

  function renderResults(data) {
    const total = data.row_count != null ? data.row_count : data.rows.length;
    resultCount.textContent = "共 " + total + " 行" +
      (data.truncated ? "（数据较多，下表仅显示前 " + data.rows.length + " 行，完整数据请导出 Excel）" : "");

    if (!data.rows.length) {
      resultSection.classList.remove("hidden");
      resultTable.innerHTML = '<div class="empty alert-error">未筛选到任何记录，请重新输入。</div>';
      totalsSection.classList.add("hidden");
      return;
    }

    let html = "<table class='data'><thead><tr>" +
      "<th>项目名称</th><th>code</th><th>staffname</th><th>manager name</th>" +
      "<th class='num'>hours</th><th>期间</th><th>描述</th></tr></thead><tbody>";
    data.rows.forEach((r) => {
      html += "<tr>" +
        "<td>" + esc(r.project_name) + "</td>" +
        "<td>" + esc(r.project_code) + "</td>" +
        "<td>" + esc(r.staff) + "</td>" +
        "<td>" + esc(r.manager) + "</td>" +
        "<td class='num'>" + fmt(r.hours) + "</td>" +
        "<td>" + esc(r.period) + "</td>" +
        "<td>" + esc(r.description) + "</td>" +
        "</tr>";
    });
    html += "</tbody></table>";
    resultTable.innerHTML = html;
    resultSection.classList.remove("hidden");

    let th = "<table class='data'><thead><tr>" +
      "<th>项目名称</th><th>code</th><th class='num'>合计工时</th>" +
      "<th class='num'>预算 (Outlook Hours)</th><th>是否超预算</th></tr></thead><tbody>";
    data.totals.forEach((t) => {
      const over = t.over ? " over" : "";
      let flag = "—";
      if (t.budget !== null && t.budget !== undefined) flag = t.over ? "超预算" : "未超";
      th += "<tr class='" + over + "'>" +
        "<td>" + esc(t.project_name) + "</td>" +
        "<td>" + esc(t.project_code) + "</td>" +
        "<td class='num'>" + fmt(t.hours) + "</td>" +
        "<td class='num'>" + fmt(t.budget) + "</td>" +
        "<td class='over-cell'>" + flag + "</td>" +
        "</tr>";
    });
    th += "</tbody></table>";
    totalsTable.innerHTML = th;
    totalsSection.classList.remove("hidden");

    clearAlertsOutside();
  }

  function clearAlertsOutside() {
    document.querySelectorAll(".page-alert").forEach((n) => n.remove());
  }

  function showBanner(msg, kind) {
    clearAlertsOutside();
    const box = document.createElement("div");
    box.className = "page-alert " + (kind === "error" ? "alert-error" : "alert-ok");
    box.textContent = msg;
    const card = document.querySelector(".section .card");
    card.parentElement.insertBefore(box, card);
    setTimeout(() => box.remove(), 4000);
  }

  async function loadStatus() {
    try {
      const data = await api("/api/timesheet/status");
      if (data.ts_rows > 0) {
        statusLine.textContent = "已缓存 " + data.ts_rows + " 条 Timesheet 记录，" + data.budgets + " 个项目预算";
      } else {
        statusLine.textContent = "尚未导入 Timesheet，请到“数据源设置”指定文件并刷新";
      }
    } catch (e) { /* ignore */ }
  }

  document.getElementById("btn-filter").addEventListener("click", async () => {
    const mode = currentMode();
    const managers = valuesOf(managerInputs);
    const codes = valuesOf(codeInputs);
    if (mode === "project" && !codes.length) {
      showBanner("请至少输入一个项目 code", "error");
      return;
    }
    if (!managers.length) {
      showBanner("请至少输入一个 manager", "error");
      return;
    }
    lastQuery = { mode: mode, codes: codes, managers: managers };
    try {
      const data = await api("/api/timesheet/filter", "POST", lastQuery);
      renderResults(data);
      if (!data.rows.length) showBanner("未筛选到任何记录，请重新输入", "error");
    } catch (err) {
      showBanner(err.message, "error");
    }
  });

  document.getElementById("btn-export").addEventListener("click", async () => {
    if (!lastQuery) {
      showBanner("请先筛选，再导出", "error");
      return;
    }
    try {
      const res = await fetch("/api/timesheet/export", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(lastQuery),
      });
      if (!res.ok) {
        let msg = "导出失败";
        try { msg = (await res.json()).error || msg; } catch (e) { /* ignore */ }
        showBanner(msg, "error");
        return;
      }
      const blob = await res.blob();
      const cd = res.headers.get("Content-Disposition") || "";
      let name = "工时筛选.xlsx";
      const m = cd.match(/filename\*=UTF-8''([^;]+)/);
      if (m) name = decodeURIComponent(m[1]);
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = name;
      a.click();
      URL.revokeObjectURL(a.href);
      showBanner("已导出：" + name, "ok");
    } catch (err) {
      showBanner(err.message, "error");
    }
  });

  document.getElementById("btn-refresh-ts").addEventListener("click", async () => {
    const btn = document.getElementById("btn-refresh-ts");
    btn.disabled = true;
    try {
      const data = await api("/api/timesheet/refresh", "POST");
      showBanner("Timesheet 已刷新：" + (data.ts_rows || 0) + " 条记录", "ok");
      await loadStatus();
    } catch (err) { showBanner(err.message, "error"); } finally { btn.disabled = false; }
  });

  document.getElementById("btn-refresh-budget").addEventListener("click", async () => {
    const btn = document.getElementById("btn-refresh-budget");
    btn.disabled = true;
    try {
      const data = await api("/api/timesheet/refresh_budget", "POST");
      showBanner("预算已刷新：" + (data.budgets || 0) + " 个项目", "ok");
      await loadStatus();
    } catch (err) { showBanner(err.message, "error"); } finally { btn.disabled = false; }
  });

  loadStatus();
})();
