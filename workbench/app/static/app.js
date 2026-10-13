(function () {
  "use strict";

  const todayEl = document.getElementById("today");
  if (todayEl) {
    const d = new Date();
    const wk = ["日", "一", "二", "三", "四", "五", "六"][d.getDay()];
    todayEl.textContent =
      d.getFullYear() + "-" +
      String(d.getMonth() + 1).padStart(2, "0") + "-" +
      String(d.getDate()).padStart(2, "0") + " 星期" + wk;
  }

  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
  }

  async function api(url, method, body) {
    const opts = { method: method || "GET", headers: {} };
    if (body !== undefined) {
      opts.headers["Content-Type"] = "application/json";
      opts.body = JSON.stringify(body);
    }
    const res = await fetch(url, opts);
    let data = {};
    try { data = await res.json(); } catch (e) { /* ignore */ }
    if (!res.ok || data.ok === false) {
      throw new Error(data.error || ("请求失败 " + res.status));
    }
    return data;
  }

  window.WB = { esc: esc, api: api };
})();
