/* =========================================================
   DÙNG CHUNG CHO 3 TRANG: MƯỢN – TRẢ, NHẬT KÝ SỰ CỐ, VẬT TƯ
   (cùng cách làm với các trang hiện có: chuông thông báo, xuất Excel...)
========================================================= */

function escapeHtml(text) {
  return String(text ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function escapeJs(text) {
  return String(text ?? "")
    .replace(/\\/g, "\\\\")
    .replace(/'/g, "\\'")
    .replace(/"/g, "&quot;")
    .replace(/\n/g, " ");
}

function normalizeSearch(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

function todayISO() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

/* "10/11/2026 13:05:00" | "10/11/2026" | "2026-11-10" -> "2026-11-10" */
function toISODate(text) {
  const value = String(text || "").trim();
  let m = value.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  m = value.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}`;
  return "";
}

function isoToVN(iso) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
}

async function readJson(response) {
  const json = await response.json().catch(() => ({}));

  if (!response.ok || !json.success) {
    throw new Error(json.error || `Lỗi máy chủ (HTTP ${response.status}).`);
  }

  return json;
}

async function postJson(url, payload, method) {
  return readJson(await fetch(url, {
    method: method || "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  }));
}

function showBox(id, text, type) {
  const box = document.getElementById(id);
  if (!box) return;
  box.className = text ? `message show ${type}` : "message";
  box.textContent = text || "";
}

function setText(id, value) {
  const el = document.getElementById(id);
  if (el) el.textContent = value;
}

function setHtml(id, value) {
  const el = document.getElementById(id);
  if (el) el.innerHTML = value;
}

/* ---------- Hộp thoại ---------- */

function openModal(id) {
  const modal = document.getElementById(id);
  if (!modal) return;
  modal.hidden = false;
  document.body.classList.add("modal-open");
  const first = modal.querySelector("input:not([type=hidden]):not([type=radio]):not([type=checkbox]), select, textarea");
  if (first) setTimeout(() => first.focus(), 30);
}

function closeModal(id) {
  const modal = document.getElementById(id);
  if (!modal) return;
  modal.hidden = true;
  if (!document.querySelector(".vlu-modal:not([hidden])")) document.body.classList.remove("modal-open");
}

document.addEventListener("keydown", event => {
  if (event.key !== "Escape") return;
  const open = [...document.querySelectorAll(".vlu-modal:not(.vlu-dialog):not([hidden])")].pop();
  if (open) closeModal(open.id);
});

/* ---------- Xuất Excel (.xls dạng bảng HTML, giống các trang hiện có) ---------- */

function exportExcel(fileName, headers, rows) {
  if (!rows.length) {
    alert("Không có dữ liệu để xuất Excel.");
    return;
  }

  const cell = value => String(value ?? "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

  const html = `
    <html><head><meta charset="UTF-8">
      <style>
        table { border-collapse: collapse; font-family: Arial, sans-serif; font-size: 12px; }
        th { background: #eaf1f8; font-weight: bold; text-align: left; }
        th, td { border: 1px solid #999999; padding: 6px; mso-number-format: "\\@"; }
      </style>
    </head><body><table>
      <thead><tr>${headers.map(h => `<th>${cell(h)}</th>`).join("")}</tr></thead>
      <tbody>${rows.map(row => `<tr>${row.map(v => `<td>${cell(v)}</td>`).join("")}</tr>`).join("")}</tbody>
    </table></body></html>`;

  const blob = new Blob(["﻿" + html], { type: "application/vnd.ms-excel;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${fileName}_${todayISO()}.xls`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/* ---------- Chuông thông báo trên header ---------- */

function getNotificationLevelClass(level) {
  if (level === "danger") return "notification-danger";
  if (level === "warning") return "notification-warning";
  return "notification-info";
}

async function loadSystemNotifications() {
  const countEl = document.getElementById("notificationCount");
  const listEl = document.getElementById("notificationList");
  const statusEl = document.getElementById("notificationStatus");

  if (!countEl || !listEl) return;

  try {
    const json = await readJson(await fetch("/api/notifications"));
    const data = json.data || [];

    countEl.innerText = json.count || 0;
    if (statusEl) statusEl.innerText = `${json.count || 0} thông báo`;

    listEl.innerHTML = data.length
      ? data.map(item => `
          <div class="notification-item ${getNotificationLevelClass(item.level)}" ${item.link ? `data-link="${escapeHtml(item.link)}"` : ""}>
            <b>${escapeHtml(item.title || "Thông báo")}</b>
            <p>${escapeHtml(item.message || "")}</p>
            <small>${escapeHtml(item.id_key || "")}</small>
          </div>`).join("")
      : `<p class="notification-empty">Chưa có thông báo mới</p>`;

  } catch (error) {
    countEl.innerText = "0";
    listEl.innerHTML = `<p class="notification-empty">Không thể tải thông báo</p>`;
    if (statusEl) statusEl.innerText = "Lỗi";
  }
}

function toggleNotificationPanel(event) {
  if (event) event.stopPropagation();
  const panel = document.getElementById("notificationPanel");
  if (panel) panel.classList.toggle("show");
}

document.addEventListener("click", event => {
  const wrapper = document.querySelector(".bell-wrapper");
  const panel = document.getElementById("notificationPanel");
  if (wrapper && panel && !wrapper.contains(event.target)) panel.classList.remove("show");
});

loadSystemNotifications();
setInterval(loadSystemNotifications, 60000);
