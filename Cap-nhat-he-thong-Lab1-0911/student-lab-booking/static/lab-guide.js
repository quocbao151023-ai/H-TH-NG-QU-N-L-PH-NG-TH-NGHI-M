/* =========================================================
   THÔNG TIN PHÒNG LAB CHO TRANG "HƯỚNG DẪN"
   ---------------------------------------------------------
   Kỹ thuật viên cập nhật nội dung tại đây (không cần sửa code khác):
     - map        : ảnh sơ đồ bố trí, đặt trong static/images/
     - mapFull    : (tuỳ chọn) bản sơ đồ có chú thích để mở / tải về
     - description: mô tả ngắn về phòng
     - areas      : mã bàn / máy sinh viên chọn nhanh ở ô "Khu vực sử dụng"
     - zones      : các khu chức năng trên sơ đồ
                    { code, name, detail, color, student: true/false }
   Lab 1 cập nhật theo Sơ đồ 5S Lab 1 (11/2026) và Bộ SOP Lab 1 v2.0.
   Ảnh sơ đồ Lab 1: xuất trang 1 của thiết kế "Sơ đồ Phòng Lab 1" trên Canva (JPG)
   rồi chép đè vào static/images/so-do-lab-1.jpg (dùng cho cả ảnh nhỏ và ảnh phóng to).
========================================================= */

window.LAB_GUIDE = {
  labs: [
    {
      name: "Lab 1",
      map: "images/so-do-lab-1.jpg",
      mapFull: "images/so-do-lab-1.jpg",
      description: "Phòng triển khai 5S, gồm 6 khu chức năng (A – F). Khi đăng ký, ghi mã bàn hoặc máy theo sơ đồ ở ô “Khu vực sử dụng”, ví dụ B3, B4 hoặc C2. Mã được dán ở góc bàn / góc máy phía lối đi.",
      areas: ["B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8", "B9", "B10", "C1", "C2", "C3", "C4", "C5", "C6", "C7"],
      zones: [
        { code: "A", name: "Khu giảng viên & quản lý Lab", detail: "Bàn giáo viên, khu quản lý Lab", color: "#2e8b3e", student: false },
        { code: "B", name: "Khu làm việc sinh viên", detail: "10 bàn B1 – B10, chia 5 cụm (mỗi cụm 2 bàn, mỗi bàn 2 ghế). Đánh số từ phía bàn giảng viên", color: "#e0a800", student: true },
        { code: "C", name: "Khu máy thực hành", detail: "C1 máy in 3D · C2 phay CNC mini · C3 tiện mini · C4 cưa lọng bàn · C5 máy cắt góc · C6 khoan bàn (dọc tường trái) · C7 máy cắt. Làm theo tờ hướng dẫn tại máy, đủ bảo hộ", color: "#d6263b", student: true },
        { code: "D", name: "Khu lưu trữ vật tư", detail: "D1 tủ T1 · D2 tủ T2 · D3 dãy kệ (kệ K1 – K3) · D4 vật liệu tấm & dài. Món dưới 1 triệu có nhãn TỰ LẤY; món nhãn đỏ BÁO KTV thì nhờ KTV lấy giúp", color: "#1d3c8f", student: false },
        { code: "E", name: "Khu ghế xếp", detail: "Ghế bổ sung cho sinh viên; dùng xong gấp gọn về khu E", color: "#6b7280", student: true },
        { code: "F", name: "Khu vệ sinh", detail: "F1 thùng rác, chổi, đồ hốt rác (cạnh máy cắt C7) · F2 máy hút bụi cỡ lớn (giữa tủ T2 và kệ K3) – báo KTV khi cần dùng", color: "#b07018", student: true }
      ]
    },
    {
      name: "Lab 2",
      map: "images/so-do-lab-2.jpg",
      description: "Phòng làm việc sinh viên gồm 4 khu (A – D) và 8 cụm bàn B1 – B8, đánh số từ cửa ra vào. Khi đăng ký, ghi mã cụm bàn ở ô “Khu vực sử dụng”, ví dụ B1 hoặc B5.",
      areas: ["B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8"],
      zones: [
        { code: "A", name: "Khu giảng viên & quản lý Lab", detail: "Bàn giáo viên, máy tính và máy in quản lý Lab", color: "#c8102e", student: false },
        { code: "B", name: "Khu làm việc sinh viên", detail: "8 cụm bàn B1 – B8 (mỗi cụm 2 bàn quay lưng, có vách ngăn và 2 ghế). B1 – B4 dãy phải, B5 – B8 dãy trái", color: "#16803d", student: true },
        { code: "C", name: "Khu tủ vật tư & máy in", detail: "Tủ vật tư 1, Tủ vật tư 2 và máy in dọc tường trên", color: "#b07018", student: false },
        { code: "D", name: "Khu trình chiếu", detail: "Bảng trắng và màn chiếu – sinh viên được mượn", color: "#6d28d9", student: true }
      ]
    },
    {
      name: "Lab 3",
      map: "images/so-do-lab-3.jpg",
      description: "Lab tự động gồm 5 khu (A – E) và 6 bàn thực hành B1 – B6, đánh số từ bàn giáo viên. Khi đăng ký, ghi mã bàn ở ô “Khu vực sử dụng”, ví dụ B1 hoặc B4.",
      areas: ["B1", "B2", "B3", "B4", "B5", "B6"],
      zones: [
        { code: "A", name: "Khu giảng viên & quản lý Lab", detail: "Bàn giáo viên, máy tính và máy in quản lý Lab", color: "#c8102e", student: false },
        { code: "B", name: "Khu bàn thực hành", detail: "6 bàn thực hành có bảng treo dụng cụ: B1 – B3 dọc tường trái, B4 – B6 dọc tường trên", color: "#16803d", student: true },
        { code: "C", name: "Khu tủ vật tư", detail: "Tủ vật tư 1 và Tủ vật tư 2 dọc tường phải", color: "#b07018", student: false },
        { code: "D", name: "Khu bảng & màn chiếu", detail: "2 bảng trắng và màn chiếu trên tường dưới", color: "#6d28d9", student: true },
        { code: "E", name: "Khu robot & thiết bị tự động", detail: "Robot Logi 01, máy in socola, robot di động; bàn cánh tay robot kèm máy nén khí", color: "#0063b1", student: false }
      ]
    }
  ]
};


/* =========================================================
   CHẾ ĐỘ HIỂN THỊ PHÒNG LAB (cổng sinh viên)
   - Mặc định chỉ hiện Lab 1: lịch, thống kê, ô chọn phòng, sơ đồ hướng dẫn.
   - Nút "Xem thêm Lab 2, Lab 3" ở banner đầu trang để mở lại tất cả.
   - Lựa chọn lưu trên trình duyệt của người xem (localStorage).
   - Máy chủ vẫn nhận đăng ký cả 3 phòng như cũ (không đổi dữ liệu).
   Đổi phòng mặc định: sửa FOCUS_LAB. Tắt hẳn chế độ này: SCOPE_ENABLED = false.
========================================================= */
(function () {
  const SCOPE_ENABLED = true;
  const FOCUS_LAB = "Lab 1";
  const KEY = "labScopeAll";

  let showAll = !SCOPE_ENABLED;
  try { showAll = showAll || localStorage.getItem(KEY) === "1"; } catch (e) { /* bỏ qua */ }
  window.LAB_SCOPE = { focus: FOCUS_LAB, showAll };

  // Sơ đồ & khu chức năng trên trang Hướng dẫn
  if (!showAll) {
    window.LAB_GUIDE.labs = window.LAB_GUIDE.labs.filter(lab => lab.name === FOCUS_LAB);
  }

  function toggle() {
    try { localStorage.setItem(KEY, showAll ? "0" : "1"); } catch (e) { /* bỏ qua */ }
    location.reload();
  }

  function addScopeBar() {
    const hero = document.querySelector(".portal-hero .container, .page-hero .container");
    if (!hero || !SCOPE_ENABLED || document.getElementById("labScopeBar")) return;
    const bar = document.createElement("div");
    bar.id = "labScopeBar";
    bar.className = "lab-scope-bar";
    bar.innerHTML = showAll
      ? `<span>Đang hiển thị: <b>Lab 1, Lab 2, Lab 3</b></span><button type="button">Chỉ hiện ${FOCUS_LAB}</button>`
      : `<span>Đang hiển thị: <b>${FOCUS_LAB}</b></span><button type="button">Xem thêm Lab 2, Lab 3</button>`;
    bar.querySelector("button").addEventListener("click", toggle);
    hero.appendChild(bar);

    const style = document.createElement("style");
    style.textContent = `
      .lab-scope-bar{display:flex;flex-wrap:wrap;align-items:center;gap:10px;margin-top:14px;font-size:14px}
      .lab-scope-bar button{border:1px solid currentColor;background:transparent;color:inherit;border-radius:999px;
        padding:6px 14px;font:inherit;font-weight:600;cursor:pointer}
      .lab-scope-bar button:hover{background:rgba(255,255,255,.15)}`;
    document.head.appendChild(style);
  }

  function keepFocusOnly(select, keepAllOption) {
    if (!select) return;
    [...select.options].forEach(option => {
      if (option.value === FOCUS_LAB) {
        option.selected = true;
        option.defaultSelected = true;      // nút "Nhập lại" vẫn giữ Lab 1
      } else if (option.value && (option.value !== "all" || !keepAllOption)) {
        option.remove();
      }
    });
    select.value = FOCUS_LAB;
  }

  function applyScope() {
    if (showAll) return;

    // Lịch: chỉ hiện Lab 1 (app.js vẽ theo ô lọc phòng)
    keepFocusOnly(document.getElementById("labFilter"), false);

    // Phiếu đăng ký: phòng cố định Lab 1
    const room = document.getElementById("room");
    if (room) {
      [...room.options].forEach(o => { if (!o.value) o.remove(); });
      keepFocusOnly(room, false);
      room.dispatchEvent(new Event("change", { bubbles: true }));
    }

    // Thanh % sử dụng và chú thích màu của Lab 2, Lab 3
    ["lab2", "lab3"].forEach(id => {
      document.getElementById(`${id}Bar`)?.closest(".progress-item")?.setAttribute("hidden", "");
      document.querySelector(`.calendar-legend .legend-box.${id}`)?.parentElement?.setAttribute("hidden", "");
    });
  }

  addScopeBar();
  applyScope();
  document.addEventListener("DOMContentLoaded", () => { addScopeBar(); applyScope(); });
})();
