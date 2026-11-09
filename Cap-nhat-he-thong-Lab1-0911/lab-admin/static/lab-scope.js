/* =========================================================
   CHẾ ĐỘ HIỂN THỊ PHÒNG LAB (Hệ thống quản lý – Dashboard KTV)
   - Mặc định chỉ hiện Lab 1 (phòng triển khai 5S của đề tài).
   - Nút "Xem thêm Lab 2, Lab 3" ở đầu mỗi trang để mở lại toàn bộ.
   - Lựa chọn lưu trên trình duyệt (localStorage), không đổi dữ liệu Google Sheet.
   - Cách lọc: dữ liệu API có trường lab / mã đăng ký L2…, L3… sẽ bị ẩn;
     ô chọn phòng chỉ còn Lab 1.
   - Thêm 3 mục menu: Mượn – trả, Sự cố, Vật tư (trang mới của lab_ops.py).
   Đổi phòng mặc định: sửa FOCUS_LAB.
========================================================= */
(function () {
  const FOCUS_LAB = "Lab 1";
  const KEY = "labScopeAll";

  function readAll() {
    try { return localStorage.getItem(KEY) === "1"; } catch (e) { return false; }
  }
  function saveAll(on) {
    try { localStorage.setItem(KEY, on ? "1" : "0"); } catch (e) { /* trình duyệt chặn lưu: chỉ áp dụng trong trang này */ }
  }

  const showAll = readAll();
  window.LAB_SCOPE = { focus: FOCUS_LAB, showAll };

  const norm = v => String(v ?? "").trim().toLowerCase().replace(/\s+/g, " ");
  const focusNorm = norm(FOCUS_LAB);
  const focusNum = (FOCUS_LAB.match(/\d+/) || [""])[0];

  function hiddenLab(item) {
    if (!item || typeof item !== "object") return false;
    const lab = item.lab ?? item.Lab ?? item.room ?? item.PHONG;
    if (lab !== undefined && lab !== null && String(lab).trim()) {
      const l = norm(lab);
      if (l.includes("tất cả") || l.includes("tat ca")) return false;
      if (/^lab ?\d$/.test(l)) return l.replace(" ", "") !== focusNorm.replace(" ", "");
    }
    const id = item.id_key ?? item.idKey ?? item.booking_id;
    const m = String(id || "").trim().toUpperCase().match(/^L([1-9])/);
    if (m) return m[1] !== focusNum;
    return false;
  }

  function filterBody(body) {
    if (Array.isArray(body)) return body.filter(x => !hiddenLab(x));
    if (body && Array.isArray(body.data)) {
      const data = body.data.filter(x => !hiddenLab(x));
      const out = Object.assign({}, body, { data });
      if (typeof body.count === "number") out.count = data.length;
      return out;
    }
    return body;
  }

  // ---------- Lọc dữ liệu API (cài trước khi các trang gọi fetch) ----------
  if (!showAll && window.fetch) {
    const original = window.fetch.bind(window);
    window.fetch = async function (input, init) {
      const response = await original(input, init);
      try {
        const url = typeof input === "string" ? input : (input && input.url) || "";
        const method = String((init && init.method) || (input && input.method) || "GET").toUpperCase();
        const type = response.headers.get("content-type") || "";
        if (method !== "GET" || !/\/api\//.test(url) || !type.includes("application/json")) return response;
        const body = await response.clone().json();
        const headers = new Headers(response.headers);
        headers.delete("content-length");
        headers.delete("content-encoding");
        return new Response(JSON.stringify(filterBody(body)), {
          status: response.status, statusText: response.statusText, headers
        });
      } catch (e) {
        return response;
      }
    };
  }

  // ---------- Giao diện ----------
  function pruneDom(root) {
    if (showAll) return;
    root.querySelectorAll("select").forEach(select => {
      [...select.options].forEach(option => {
        const t = norm(option.value || option.textContent);
        if (/^lab ?\d$/.test(t) && t.replace(" ", "") !== focusNorm.replace(" ", "")) option.remove();
      });
    });
    root.querySelectorAll("button, label, li, .legend-item, .chip, .lab-chip, span").forEach(el => {
      if (el.children.length > 2) return;
      const t = norm(el.textContent);
      if (/^lab ?[2-9]$/.test(t)) el.style.display = "none";
    });
  }

  function addScopeBar() {
    const hero = document.querySelector(".page-hero .container");
    if (!hero || document.getElementById("labScopeBar")) return;
    const bar = document.createElement("div");
    bar.id = "labScopeBar";
    bar.className = "lab-scope-bar";
    bar.innerHTML = showAll
      ? `<span>Đang hiển thị: <b>tất cả phòng Lab</b></span><button type="button">Chỉ hiện ${FOCUS_LAB}</button>`
      : `<span>Đang hiển thị: <b>${FOCUS_LAB}</b> (phòng triển khai 5S)</span><button type="button">Xem thêm Lab 2, Lab 3</button>`;
    bar.querySelector("button").addEventListener("click", () => { saveAll(!showAll); location.reload(); });
    hero.appendChild(bar);
  }

  function addNavLinks() {
    if (!document.body.classList.contains("role-admin")) return;
    const nav = document.querySelector(".main-nav");
    if (!nav || nav.querySelector('a[href="/loans"]')) return;
    const anchor = nav.querySelector('a[href="/attendance"]') || nav.lastElementChild;
    const links = [
      ["/loans", "Mượn – trả", '<path fill-rule="evenodd" clip-rule="evenodd" d="M7 2h10v1.5h3V22H4V3.5h3V2Zm1.5 1.5v1h7v-1h-7ZM5.5 5v15.5h13V5H17v1.5H7V5H5.5Zm3 5h7v1.5h-7V10Zm0 3.5h7V15h-7v-1.5Zm0 3.5H13v1.5H8.5V17Z" fill="currentColor"/>'],
      ["/incidents", "Sự cố", '<path fill-rule="evenodd" clip-rule="evenodd" d="M2.86 21c-.45 0-.855-.235-1.08-.625a1.23 1.23 0 0 1 0-1.25L10.915 3.14a1.25 1.25 0 0 1 2.17 0l9.14 15.99c.225.39.22.86 0 1.25-.225.39-.63.625-1.08.625L21.14 21H2.86Zm.435-1.5H20.71L12 4.26 3.29 19.5h.005Zm9.455-5V9h-1.5v5.5h1.5ZM13 17a1 1 0 1 1-2 0 1 1 0 0 1 2 0Z" fill="currentColor"/>'],
      ["/inventory", "Vật tư", '<path fill-rule="evenodd" clip-rule="evenodd" d="M12 2 3 6.5v11L12 22l9-4.5v-11L12 2Zm6.82 4.75L12 10.16 5.18 6.75 12 3.34l6.82 3.41ZM4.5 7.96l6.75 3.38v8.82L4.5 16.79V7.96Zm8.25 12.2v-8.82l6.75-3.38v8.83l-6.75 3.37Z" fill="currentColor"/>']
    ];
    let after = anchor;
    links.forEach(([href, text, icon]) => {
      const a = document.createElement("a");
      a.href = href;
      a.innerHTML = `<svg class="gi" viewBox="0 0 24 24" aria-hidden="true">${icon}</svg>${text}`;
      if (location.pathname === href) {
        nav.querySelectorAll("a.active").forEach(x => x.classList.remove("active"));
        a.classList.add("active");
      }
      after.insertAdjacentElement("afterend", a);
      after = a;
    });
  }

  function addStyle() {
    const css = `
      .lab-scope-bar{display:flex;flex-wrap:wrap;align-items:center;gap:10px;margin-top:14px;font-size:14px}
      .lab-scope-bar span{opacity:.92}
      .lab-scope-bar button{border:1px solid currentColor;background:transparent;color:inherit;border-radius:999px;
        padding:6px 14px;font:inherit;font-weight:600;cursor:pointer}
      .lab-scope-bar button:hover{background:rgba(255,255,255,.15)}
      /* Menu kỹ thuật viên có thêm 3 mục (Mượn – trả, Sự cố, Vật tư): thu gọn để không tràn */
      @media (min-width:1081px){
        .role-admin .main-nav{gap:0;min-width:0}
        .role-admin .main-nav a{padding:0 9px;font-size:12px;letter-spacing:.1px;gap:6px}
        .role-admin .main-nav a::after{left:9px;right:9px}
        .role-admin .main-nav a .gi{display:none}
      }
      /* Màn hình 1081–1340px: đưa menu xuống hàng riêng (giống bố cục máy tính bảng của theme) */
      @media (min-width:1081px) and (max-width:1340px){
        .role-admin .site-header .container{height:auto;flex-wrap:wrap;gap:0 16px;padding-top:10px}
        .role-admin .brand{flex:1}
        .role-admin .main-nav{order:3;flex:0 0 100%;height:48px;justify-content:flex-start;overflow-x:auto;border-top:1px solid var(--line-soft)}
      }`;
    const style = document.createElement("style");
    style.textContent = css;
    document.head.appendChild(style);
  }

  document.addEventListener("DOMContentLoaded", () => {
    addStyle();
    addNavLinks();
    addScopeBar();
    pruneDom(document);
    // các ô chọn / nút do script trang tạo sau khi tải dữ liệu
    setTimeout(() => pruneDom(document), 800);
    setTimeout(() => pruneDom(document), 2500);
  });
})();
