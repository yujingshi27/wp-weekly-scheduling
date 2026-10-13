(function () {
  "use strict";
  const { esc, api } = window.WB;

  const tabsEl = document.getElementById("todo-tabs");
  if (!tabsEl) return;

  const listEl = document.getElementById("todo-list");
  const archiveBar = document.getElementById("archive-bar");
  const archiveSearch = document.getElementById("archive-search");
  const mask = document.getElementById("modal-mask");
  const form = document.getElementById("todo-form");
  const repeatBox = document.getElementById("repeat-box");

  let currentTab = "今日得做";
  let editingId = null;
  let selectedTags = [];
  let allTags = [];

  const STATUS_OPTIONS = ["待完成", "今日得做", "卡着", "等反馈"];

  function catBadge(cat) {
    const cls = cat === "DDL" ? "blue" : cat === "常规重复" ? "green" : "";
    return '<span class="badge ' + cls + '">' + esc(cat) + "</span>";
  }

  function quadBadge(q) {
    if (!q) return "";
    const cls = { "重要且紧急": "red", "重要不紧急": "amber", "紧急不重要": "blue", "不紧急不重要": "gray" }[q] || "";
    return '<span class="badge ' + cls + '">' + esc(q) + "</span>";
  }

  function repeatLabel(item) {
    if (item.category !== "常规重复" || !item.repeat_type || item.repeat_type === "none") return "";
    const names = { daily: "每天", weekly: "每周", monthly: "每月", workday: "工作日" };
    let extra = "";
    try {
      const p = item.repeat_param ? JSON.parse(item.repeat_param) : {};
      if (item.repeat_type === "weekly" && p.weekdays) {
        extra = "（周" + p.weekdays.map((n) => "一二三四五六日"[n]).join("、") + "）";
      } else if (item.repeat_type === "monthly" && p.day) {
        extra = "（" + p.day + "号）";
      }
    } catch (e) { /* ignore */ }
    return '<span class="badge">' + (names[item.repeat_type] || item.repeat_type) + esc(extra) + "</span>";
  }

  function todoRow(item) {
    const meta = [];
    meta.push(catBadge(item.category));
    const qb = quadBadge(item.quadrant);
    if (qb) meta.push(qb);
    if (item.due_date) {
      meta.push('<span class="' + (item.overdue ? "text-danger" : "") + '">截止 ' + esc(item.due_date) + "</span>");
    }
    if (item.overdue) meta.push('<span class="badge red">已逾期</span>');
    const rl = repeatLabel(item);
    if (rl) meta.push(rl);
    (item.tags || []).forEach((t) => meta.push('<span class="badge amber">' + esc(t) + "</span>"));

    const statusOpts = STATUS_OPTIONS.map(
      (s) => '<option value="' + s + '"' + (s === item.status ? " selected" : "") + ">" + s + "</option>"
    ).join("");
    const note = item.note ? '<div class="todo-note">' + esc(item.note) + "</div>" : "";

    return (
      '<div class="todo-row' + (item.overdue ? " overdue" : "") + '" data-id="' + item.id + '">' +
        '<div class="todo-main">' +
          '<div class="todo-title">' + esc(item.title) + "</div>" +
          '<div class="todo-meta">' + meta.join('<span class="dot">·</span>') + "</div>" +
          note +
        "</div>" +
        '<div class="todo-actions">' +
          '<select class="row-status">' + statusOpts + "</select>" +
          '<button class="btn primary btn-sm act-complete">完成</button>' +
          '<button class="btn btn-sm act-edit">编辑</button>' +
          '<button class="btn danger btn-sm act-delete">删除</button>' +
        "</div>" +
      "</div>"
    );
  }

  function archiveRow(item) {
    const meta = [catBadge(item.category)];
    const qb = quadBadge(item.quadrant);
    if (qb) meta.push(qb);
    if (item.due_date) meta.push("<span>截止 " + esc(item.due_date) + "</span>");
    if (item.archived_at) meta.push("<span>归档 " + esc(String(item.archived_at).slice(0, 10)) + "</span>");
    (item.tags || []).forEach((t) => meta.push('<span class="badge amber">' + esc(t) + "</span>"));
    return (
      '<div class="todo-row" data-id="' + item.id + '">' +
        '<div class="todo-main">' +
          '<div class="todo-title">' + esc(item.title) + "</div>" +
          '<div class="todo-meta">' + meta.join('<span class="dot">·</span>') + "</div>" +
          (item.note ? '<div class="todo-note">' + esc(item.note) + "</div>" : "") +
        "</div>" +
        '<div class="todo-actions">' +
          '<button class="btn danger btn-sm act-delete">删除</button>' +
        "</div>" +
      "</div>"
    );
  }

  async function loadTodos() {
    const data = await api("/api/todos");
    Object.keys(data.counts).forEach((k) => {
      const el = tabsEl.querySelector('[data-count="' + k + '"]');
      if (el) el.textContent = data.counts[k];
    });
    if (currentTab === "归档") return;
    const items = data.groups[currentTab] || [];
    listEl.innerHTML = items.length ? items.map(todoRow).join("") : '<div class="empty">暂无待办</div>';
  }

  async function loadArchive() {
    const q = (archiveSearch.value || "").trim();
    const data = await api("/api/todos/archive" + (q ? "?q=" + encodeURIComponent(q) : ""));
    listEl.innerHTML = data.items.length ? data.items.map(archiveRow).join("") : '<div class="empty">归档区为空</div>';
  }

  function refreshCurrent() {
    return currentTab === "归档" ? loadArchive() : loadTodos();
  }

  tabsEl.addEventListener("click", (e) => {
    const tab = e.target.closest(".tab");
    if (!tab) return;
    tabsEl.querySelectorAll(".tab").forEach((t) => t.classList.remove("active"));
    tab.classList.add("active");
    currentTab = tab.dataset.tab;
    archiveBar.classList.toggle("hidden", currentTab !== "归档");
    refreshCurrent().catch((err) => alert(err.message));
  });

  listEl.addEventListener("click", async (e) => {
    const row = e.target.closest(".todo-row");
    if (!row) return;
    const id = row.dataset.id;
    try {
      if (e.target.classList.contains("act-complete")) {
        await api("/api/todos/" + id + "/complete", "POST");
        await refreshCurrent();
      } else if (e.target.classList.contains("act-delete")) {
        if (confirm("确定删除这条待办？")) {
          await api("/api/todos/" + id, "DELETE");
          await refreshCurrent();
        }
      } else if (e.target.classList.contains("act-edit")) {
        const data = await api("/api/todos/" + id);
        if (data.item) openModal(data.item);
      }
    } catch (err) { alert(err.message); }
  });

  listEl.addEventListener("change", async (e) => {
    if (!e.target.classList.contains("row-status")) return;
    const row = e.target.closest(".todo-row");
    try {
      await api("/api/todos/" + row.dataset.id, "PATCH", { status: e.target.value });
      await refreshCurrent();
    } catch (err) { alert(err.message); }
  });

  let archiveTimer = null;
  archiveSearch.addEventListener("input", () => {
    clearTimeout(archiveTimer);
    archiveTimer = setTimeout(() => loadArchive().catch((e) => alert(e.message)), 250);
  });

  async function loadTags() {
    const data = await api("/api/tags");
    allTags = data.tags;
    renderTagChips();
  }

  function renderTagChips() {
    const box = document.getElementById("tag-chips");
    box.innerHTML = allTags
      .map((t) => '<span class="chip' + (selectedTags.includes(t) ? " on" : "") + '" data-tag="' + esc(t) + '">' + esc(t) + "</span>")
      .join("");
  }

  document.getElementById("tag-chips").addEventListener("click", (e) => {
    const chip = e.target.closest(".chip");
    if (!chip) return;
    const t = chip.dataset.tag;
    const i = selectedTags.indexOf(t);
    if (i >= 0) selectedTags.splice(i, 1); else selectedTags.push(t);
    renderTagChips();
  });

  document.getElementById("btn-add-tag").addEventListener("click", async () => {
    const input = document.getElementById("new-tag");
    const name = input.value.trim();
    if (!name) return;
    try {
      await api("/api/tags", "POST", { name: name });
      input.value = "";
      if (!selectedTags.includes(name)) selectedTags.push(name);
      await loadTags();
    } catch (err) { alert(err.message); }
  });

  const typeSelect = document.getElementById("f-repeat-type");
  const weeklyBox = document.getElementById("repeat-weekly");
  const monthlyBox = document.getElementById("repeat-monthly");

  function syncRepeatUI() {
    const isRepeat = document.getElementById("f-category").value === "常规重复";
    repeatBox.classList.toggle("hidden", !isRepeat);
    if (!isRepeat) return;
    weeklyBox.classList.toggle("hidden", typeSelect.value !== "weekly");
    monthlyBox.classList.toggle("hidden", typeSelect.value !== "monthly");
  }

  document.getElementById("f-category").addEventListener("change", syncRepeatUI);
  typeSelect.addEventListener("change", syncRepeatUI);

  function openModal(item) {
    editingId = item ? item.id : null;
    document.getElementById("modal-title").textContent = item ? "编辑待办" : "新建待办";
    document.getElementById("f-title").value = item ? item.title : "";
    document.getElementById("f-due").value = item && item.due_date ? item.due_date : "";
    document.getElementById("f-status").value = item ? item.status : "待完成";
    document.getElementById("f-category").value = item ? item.category : "日常任务";
    document.getElementById("f-quadrant").value = item && item.quadrant ? item.quadrant : "";
    selectedTags = item ? (item.tags || []).slice() : [];
    document.getElementById("f-note").value = item ? item.note || "" : "";
    const rt = item && item.repeat_type && item.repeat_type !== "none" ? item.repeat_type : "daily";
    typeSelect.value = rt;
    let param = {};
    try { param = item && item.repeat_param ? JSON.parse(item.repeat_param) : {}; } catch (e) { /* ignore */ }
    weeklyBox.querySelectorAll("input").forEach((c) => {
      c.checked = (param.weekdays || []).map(String).includes(c.value);
    });
    document.getElementById("f-repeat-day").value = param.day || 1;
    renderTagChips();
    syncRepeatUI();
    mask.classList.remove("hidden");
  }

  function closeModal() {
    mask.classList.add("hidden");
    editingId = null;
  }

  document.getElementById("btn-new").addEventListener("click", () => openModal(null));
  document.getElementById("modal-close").addEventListener("click", closeModal);
  document.getElementById("modal-cancel").addEventListener("click", closeModal);
  mask.addEventListener("click", (e) => { if (e.target === mask) closeModal(); });

  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    const category = document.getElementById("f-category").value;
    const payload = {
      title: document.getElementById("f-title").value.trim(),
      due_date: document.getElementById("f-due").value,
      status: document.getElementById("f-status").value,
      category: category,
      quadrant: document.getElementById("f-quadrant").value,
      note: document.getElementById("f-note").value,
      tags: selectedTags.slice(),
    };
    if (category === "常规重复") {
      payload.repeat_type = typeSelect.value;
      if (typeSelect.value === "weekly") {
        const days = [];
        weeklyBox.querySelectorAll("input:checked").forEach((c) => days.push(parseInt(c.value, 10)));
        if (!days.length) return alert("请选择每周几");
        payload.repeat_param = { weekdays: days };
      } else if (typeSelect.value === "monthly") {
        payload.repeat_param = { day: parseInt(document.getElementById("f-repeat-day").value, 10) || 1 };
      } else {
        payload.repeat_param = {};
      }
    } else {
      payload.repeat_type = "none";
      payload.repeat_param = {};
    }
    try {
      if (editingId) await api("/api/todos/" + editingId, "PATCH", payload);
      else await api("/api/todos", "POST", payload);
      closeModal();
      await loadTodos();
    } catch (err) { alert(err.message); }
  });

  document.getElementById("btn-refresh-ddl").addEventListener("click", async () => {
    try {
      const data = await api("/api/ddl/refresh", "POST");
      await loadTodos();
      alert("DDL 刷新完成：新增 " + (data.added || 0) + " 条，更新 " + (data.updated || 0) + " 条。");
    } catch (err) { alert(err.message); }
  });

  loadTags().catch(() => {});
  loadTodos().catch((err) => alert(err.message));
})();
