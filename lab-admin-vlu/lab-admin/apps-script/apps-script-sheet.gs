/**
 * =====================================================================
 * APPS SCRIPT CỦA FILE GOOGLE SHEET PHÒNG LAB (không còn dùng Google Form)
 * ---------------------------------------------------------------------
 * 2 giao diện ghi dữ liệu vào Sheet với EMAIL_STATUS = "CHƯA GỬI".
 * Script tự chạy mỗi 1 phút và GỬI EMAIL:
 *
 * Hàm chính: guiEmail (trình kích hoạt mỗi 1 phút), xử lý lần lượt:
 *  1. Email đăng ký   – sheet WEB_BOOKING (cổng sinh viên), theo thứ tự:
 *     a. Lượt bị hủy vì sinh viên ĐĂNG KÝ TRÙNG LỊCH (cùng ngày, cùng ca):
 *        gửi email "hệ thống đã hủy lượt đăng ký trước".
 *     b. Lượt mới: kiểm tra khung giờ còn chỗ (tối đa 9 nhóm / phòng / ca)
 *        - còn chỗ  -> email XÁC NHẬN (mẫu cũ, kèm file lịch .ics)
 *        - hết chỗ  -> chuyển lượt sang "TỪ CHỐI - ĐỦ CHỖ", email KHÔNG THÀNH CÔNG nhắc đăng ký lại.
 *  2. Email lịch trực – sheet LAB_STAFF_SCHEDULE (hệ thống quản lý):
 *     email phân công / hủy ca trực cho KTV, trợ giảng, giảng viên.
 *
 *  Gửi xong ghi EMAIL_STATUS = "ĐÃ GỬI" (lỗi: "LỖI" + lý do ở EMAIL_ERROR).
 *
 * CÀI ĐẶT (1 lần, bằng tài khoản kỹ thuật viên):
 *   Tiện ích mở rộng -> Apps Script -> XÓA hết code / tệp .gs cũ, dán file này, Lưu
 *   -> chọn hàm caiDatTuDong -> Chạy -> cấp quyền.
 *   Gửi ngay / kiểm tra: chọn hàm guiEmail -> Chạy, xem "Nhật ký thực thi" (ghi rõ đã gửi bao nhiêu, lỗi gì).
 * =====================================================================
 */

const CAU_HINH = {
  SPREADSHEET_ID: "16eguXNHfCzbTq6olAQaMQBAylkqNR3gjVMmtnAkSBpE",
  SHEET_DANG_KY: "WEB_BOOKING",
  SHEET_LICH_TRUC: "LAB_STAFF_SCHEDULE",
  TOI_DA_NHOM: 9,                                         // nhóm / phòng / ca
  TOI_DA_EMAIL_MOI_LAN: 30,
  TEN_DON_VI: "Khoa Kỹ thuật và Quản lý Công nghiệp",
  EMAIL_DON_VI: "k.iem@vlu.edu.vn",
  DIA_DIEM: "Trường Đại học Văn Lang - Cơ sở 2",
  BANNER_FILE_ID: "1gCKhNkiYSqE0oAcG2v_VrQlGg4XzZsJm",   // ảnh banner cuối email (Google Drive)
  GIO_CA: {
    "Ca 1": ["07:00", "09:30"],
    "Ca 2": ["09:30", "12:00"],
    "Ca 3": ["13:00", "15:30"],
    "Ca 4": ["15:30", "18:00"]
  },
  PHONG: { "Lab 1": "Lab 1 - CS2.D.1.01", "Lab 2": "Lab 2 - CS2.D.1.04", "Lab 3": "Lab 3 - CS2.D.1.05" }
};

const TT_XAC_NHAN = "ĐÃ XÁC NHẬN";
const TT_HET_CHO = "TỪ CHỐI - ĐỦ CHỖ";
const EMAIL_CHO = "CHƯA GỬI";

/* =====================================================================
   HÀM CHÍNH – trình kích hoạt gọi mỗi 1 phút (muốn gửi ngay: chọn guiEmail -> Chạy)
===================================================================== */

function guiEmail() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    Logger.log("Lần chạy trước chưa xong, bỏ qua lần này.");
    return;
  }

  try {
    [["Email đăng ký (WEB_BOOKING)", xuLyDangKy_], ["Email lịch trực (LAB_STAFF_SCHEDULE)", xuLyLichTruc_]]
      .forEach(([ten, ham]) => {
        try {
          const kq = ham();
          Logger.log(`${ten}: chờ gửi ${kq.cho}, đã gửi ${kq.gui}, lỗi ${kq.loi.length}` +
            (kq.ghiChu ? ` – ${kq.ghiChu}` : "") +
            (kq.loi.length ? `\n  Lỗi: ${kq.loi.join(" | ")}` : ""));
        } catch (err) {
          Logger.log(`${ten}: LỖI – ${(err && err.stack) || err}`);
        }
      });
  } finally {
    lock.releaseLock();
  }
}

/** Giữ tên cũ: chạy tay hàm nào cũng xử lý cả 2 loại email. */
function guiEmailDangKy() { guiEmail(); }
function guiEmailLichTruc() { guiEmail(); }

/* =====================================================================
   1. EMAIL ĐĂNG KÝ (WEB_BOOKING)
===================================================================== */

function xuLyDangKy_() {
  {
    const sent = { n: 0, gui: 0, loi: [] };
    const sheet = moSheet_(CAU_HINH.SHEET_DANG_KY);
    if (!sheet) return Object.assign(sent, { cho: 0, ghiChu: "chưa có sheet " + CAU_HINH.SHEET_DANG_KY });
    if (sheet.getLastRow() < 2) return Object.assign(sent, { cho: 0, ghiChu: "chưa có lượt đăng ký" });

    const values = sheet.getDataRange().getDisplayValues();
    const C = cot_(values[0]);
    const ghi = ghiO_(sheet, C);
    const rows = values.slice(1).map((r, i) => Object.assign(r, { __row: i + 2 }));
    if (C.EMAIL_STATUS === undefined) return Object.assign(sent, { cho: 0, ghiChu: "thiếu cột EMAIL_STATUS" });
    sent.cho = rows.filter(r => laCho_(r, C)).length;

    // a. Ca bị hủy do đăng ký trùng lịch -> báo cho sinh viên (các ca cùng mã gộp 1 email)
    const huy = {};
    rows.filter(r => laCho_(r, C) && String(r[C.BOOKING_STATUS]).indexOf("ĐÃ HỦY") === 0)
      .forEach(r => {
        const key = String(r[C.ID_KEY]).trim() + "|" + String(r[C.THAY_THE_BOI] || "").trim();
        (huy[key] = huy[key] || []).push(r);
      });
    Object.keys(huy).forEach(key => guiMot_(sent, huy[key], ghi, () => emailHuyTrung_(huy[key], C, rows)));

    // b. Lượt mới: kiểm tra còn chỗ -> xác nhận / không thành công
    const choXuLy = rows.filter(r => laCho_(r, C) && String(r[C.BOOKING_STATUS]).trim() === TT_XAC_NHAN);
    const nhom = {};
    choXuLy.forEach(r => {
      const key = String(r[C.MA_DANG_KY_CHUNG] || "").trim() || String(r[C.ID_KEY]).trim();
      (nhom[key] = nhom[key] || []).push(r);
    });

    Object.keys(nhom).forEach(key => {
      const ok = [], full = [];
      nhom[key].forEach(r => {
        const thuTu = thuTuTrongCa_(r, rows, C);
        r.__thuTu = thuTu;
        (thuTu <= CAU_HINH.TOI_DA_NHOM ? ok : full).push(r);
      });

      if (full.length) {
        full.forEach(r => {
          r[C.BOOKING_STATUS] = TT_HET_CHO;   // để các lượt sau trong cùng lần chạy tính đúng thứ tự
          ghi(r.__row, "BOOKING_STATUS", TT_HET_CHO);
          ghi(r.__row, "GHI_CHU_HE_THONG", "Khung giờ đã đủ " + CAU_HINH.TOI_DA_NHOM + " nhóm");
        });
        guiMot_(sent, full, ghi, () => emailHetCho_(full, C));
      }
      if (ok.length) guiMot_(sent, ok, ghi, () => emailXacNhan_(ok, C));
    });
    return sent;
  }
}

/** Thứ tự đăng ký của lượt trong khung giờ (theo dòng trên sheet, chỉ tính lượt còn hiệu lực). */
function thuTuTrongCa_(row, rows, C) {
  const slot = chuanSlot_(row[C.SLOT_KEY]);
  let n = 0;
  for (const r of rows) {
    if (String(r[C.BOOKING_STATUS]).trim() === TT_XAC_NHAN && chuanSlot_(r[C.SLOT_KEY]) === slot) n++;
    if (r === row) return n;
  }
  return n;
}

function emailXacNhan_(rows, C) {
  const r = rows[0];
  const nhieuCa = rows.length > 1;
  const laNhom = String(r[C.HINH_THUC] || "").toUpperCase() === "NHÓM";
  const cheDo = {
    "ĐẠI DIỆN": "Người đại diện bàn giao cho cả nhóm",
    "TỪNG THÀNH VIÊN": "Từng thành viên tự bàn giao bằng MSSV của mình"
  }[String(r[C.HINH_THUC_BAN_GIAO] || "").toUpperCase()] || "";

  // 1 mã ID_KEY dùng chung cho mọi ca đã đăng ký cùng lúc -> 1 bảng, liệt kê các ca
  const bang = bangThongTin_([
    ["ID_KEY", `<span style="font-weight:bold;color:#d32f2f">${esc_(r[C.ID_KEY])}</span>` +
      (nhieuCa ? ` <span style="color:#555">(dùng chung cho ${rows.length} ca)</span>` : "")],
    ["SLOT_KEY", rows.map(x => esc_(x[C.SLOT_KEY])).join("<br>")],
    ["Phòng Lab", esc_(CAU_HINH.PHONG[r[C.PHONG]] || r[C.PHONG])],
    ["Ngày sử dụng", esc_(ngayVN_(r[C.NGAY_SU_DUNG]))],
    ["Ca sử dụng", rows.map(x => esc_(caHienThi_(x[C.CA_SU_DUNG]))).join("<br>")],
    ["Khu vực", esc_(r[C.KHU_VUC])],
    ["Số lượng đăng ký hiện tại", nhieuCa
      ? rows.map(x => esc_(x[C.CA_SU_DUNG]) + ": " + x.__thuTu + "/" + CAU_HINH.TOI_DA_NHOM).join("<br>")
      : r.__thuTu + "/" + CAU_HINH.TOI_DA_NHOM],
    ["Hình thức", laNhom ? "Nhóm – " + esc_(r[C.MEMBER_COUNT]) + " người" : "Cá nhân"],
    ["Hình thức bàn giao", laNhom ? esc_(cheDo) : ""]
  ]);

  const anh = anhBanner_();
  return {
    to: r[C.EMAIL_NHAN_XAC_NHAN],
    subject: "XÁC NHẬN ĐĂNG KÝ SỬ DỤNG PHÒNG LAB THÀNH CÔNG",
    html: thanEmail_(`
      <p>Thân gửi bạn <b>${esc_(r[C.TEN_DAI_DIEN])}</b>,</p>
      <p>${CAU_HINH.TEN_DON_VI} xác nhận đăng ký sử dụng phòng Lab của bạn đã <b>thành công</b>.</p>
      <p><b>Thông tin lịch đăng ký${nhieuCa ? " (" + rows.length + " ca)" : ""}:</b></p>
      ${bang}
      <p style="color:#d32f2f;font-weight:bold;margin:12px 0">
        Lưu ý: Khi vào phòng, vui lòng CHỤP HÌNH TRƯỚC NGAY LẬP TỨC để làm minh chứng bàn giao.
      </p>
      <p>Sau khi sử dụng, vệ sinh khu vực, chụp hình SAU và bàn giao trên cổng đăng ký bằng ID_KEY + MSSV.
        ${nhieuCa ? "Mã này dùng chung cho các ca trên, chỉ cần <b>bàn giao 1 lần sau ca cuối</b>." : ""}</p>
      <p>Vui lòng sử dụng phòng đúng mục đích đã đăng ký, tuân thủ nội quy phòng Lab và thực hiện vệ sinh khu vực sau khi sử dụng.</p>`, anh),
    images: anh,
    attachments: [fileLich_(rows, C)]
  };
}

function emailHetCho_(rows, C) {
  const r = rows[0];
  const anh = anhBanner_();
  return {
    to: r[C.EMAIL_NHAN_XAC_NHAN],
    subject: "THÔNG BÁO HỦY ĐĂNG KÝ PHÒNG LAB – VUI LÒNG ĐĂNG KÝ LẠI",
    html: thanEmail_(`
      <p>Thân gửi bạn <b>${esc_(r[C.TEN_DAI_DIEN])}</b>,</p>
      <p>${CAU_HINH.TEN_DON_VI} rất tiếc thông báo đăng ký sử dụng phòng Lab của bạn <b>không thể được xác nhận</b>
        do khung giờ này đã đủ số lượng tiếp nhận.</p>
      <p><b>Thông tin đăng ký:</b></p>
      ${bangThongTin_([
        ["Phòng Lab", esc_(CAU_HINH.PHONG[r[C.PHONG]] || r[C.PHONG])],
        ["Ngày sử dụng", esc_(ngayVN_(r[C.NGAY_SU_DUNG]))],
        ["Ca sử dụng", rows.map(x => esc_(caHienThi_(x[C.CA_SU_DUNG]))).join("<br>")]
      ])}
      <p><b>Địa điểm:</b> ${esc_(CAU_HINH.DIA_DIEM)}</p>
      <p>Bạn vui lòng chọn <b>phòng Lab khác</b> hoặc <b>ca khác</b> trên cổng đăng ký để đăng ký lại.</p>`, anh),
    images: anh
  };
}

function emailHuyTrung_(huyRows, C, rows) {
  const r = huyRows[0];
  const maCu = String(r[C.ID_KEY]).trim();
  const maMoi = String(r[C.THAY_THE_BOI] || "").trim();
  const moi = rows.filter(x => String(x[C.ID_KEY]).trim() === maMoi && String(x[C.BOOKING_STATUS]).trim() === TT_XAC_NHAN);
  // Các ca khác của mã cũ không bị trùng thì vẫn giữ
  const conGiu = rows.filter(x => String(x[C.ID_KEY]).trim() === maCu && String(x[C.BOOKING_STATUS]).trim() === TT_XAC_NHAN);
  const ca = list => list.map(x => esc_(caHienThi_(x[C.CA_SU_DUNG]))).join("<br>");
  const anh = anhBanner_();
  return {
    to: r[C.EMAIL_NHAN_XAC_NHAN],
    subject: "THÔNG BÁO ĐĂNG KÝ TRÙNG LỊCH – HỆ THỐNG ĐÃ HỦY LƯỢT ĐĂNG KÝ TRƯỚC",
    html: thanEmail_(`
      <p>Thân gửi bạn <b>${esc_(r[C.TEN_DAI_DIEN])}</b>,</p>
      <p>Hệ thống ghi nhận bạn vừa đăng ký thêm một lượt <b>trùng ngày và ca</b> với lượt đăng ký trước đó.
        Để tránh trùng lịch, hệ thống <b>đã hủy ca trùng của lượt đăng ký trước</b> và giữ lượt đăng ký mới.</p>
      <p><b>Ca đã hủy:</b></p>
      ${bangThongTin_([
        ["ID_KEY", `<span style="font-weight:bold;color:#d32f2f;text-decoration:line-through">${esc_(maCu)}</span>`],
        ["Phòng Lab", esc_(CAU_HINH.PHONG[r[C.PHONG]] || r[C.PHONG])],
        ["Ngày sử dụng", esc_(ngayVN_(r[C.NGAY_SU_DUNG]))],
        ["Ca sử dụng", ca(huyRows)]
      ])}
      ${conGiu.length ? `<p>Các ca khác của mã <b>${esc_(maCu)}</b> vẫn giữ nguyên: ${conGiu.map(x => esc_(x[C.CA_SU_DUNG])).join(", ")}.</p>` : ""}
      <p><b>Lượt đăng ký đang giữ:</b></p>
      ${bangThongTin_([
        ["ID_KEY", `<span style="font-weight:bold;color:#2e7d32">${esc_(maMoi)}</span>`],
        ["Phòng Lab", moi.length ? esc_(CAU_HINH.PHONG[moi[0][C.PHONG]] || moi[0][C.PHONG]) : ""],
        ["Ngày sử dụng", moi.length ? esc_(ngayVN_(moi[0][C.NGAY_SU_DUNG])) : ""],
        ["Ca sử dụng", ca(moi)]
      ])}
      <p>Nếu bạn không thực hiện đăng ký mới, vui lòng liên hệ ${CAU_HINH.TEN_DON_VI} để được hỗ trợ.</p>`, anh),
    images: anh
  };
}

/* =====================================================================
   2. EMAIL LỊCH TRỰC (LAB_STAFF_SCHEDULE)
===================================================================== */

function xuLyLichTruc_() {
  {
    const sent = { n: 0, gui: 0, loi: [] };
    const sheet = moSheet_(CAU_HINH.SHEET_LICH_TRUC);
    if (!sheet) return Object.assign(sent, { cho: 0, ghiChu: "chưa có sheet " + CAU_HINH.SHEET_LICH_TRUC });
    if (sheet.getLastRow() < 2) return Object.assign(sent, { cho: 0, ghiChu: "chưa có lịch trực" });

    const values = sheet.getDataRange().getDisplayValues();
    const C = cot_(values[0]);
    const ghi = ghiO_(sheet, C);
    if (C.EMAIL_STATUS === undefined || C["Email"] === undefined) {
      return Object.assign(sent, { cho: 0, ghiChu: "thiếu cột Email / EMAIL_STATUS" });
    }
    sent.cho = values.slice(1).filter(r => laCho_(r, C)).length;

    // Các ca thêm cùng lúc cho 1 người gộp 1 email; ca hủy gửi riêng
    const nhom = {};
    values.slice(1).forEach((r, i) => {
      r.__row = i + 2;
      if (!laCho_(r, C) || !String(r[C["Email"]] || "").trim()) return;
      const huy = String(r[C["TRẠNG THÁI"]]).trim() === "ĐÃ HỦY";
      const key = [String(r[C["Email"]]).toLowerCase(), huy ? "HUY-" + i : "PC-" + r[C["Thời gian tạo"]]].join("|");
      (nhom[key] = nhom[key] || []).push(r);
    });

    Object.keys(nhom).forEach(key => guiMot_(sent, nhom[key], ghi, () => emailLichTruc_(nhom[key], C)));
    return sent;
  }
}

function emailLichTruc_(rows, C) {
  const r = rows[0];
  const huy = String(r[C["TRẠNG THÁI"]]).trim() === "ĐÃ HỦY";
  const thu = ["Chủ nhật", "Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6", "Thứ 7"];
  const td = "padding:8px 12px;border:1px solid #d9d9d9";
  const dong = rows.slice()
    .sort((a, b) => ngayDate_(a[C["Ngày"]]) - ngayDate_(b[C["Ngày"]]))
    .map(x => {
      const d = ngayDate_(x[C["Ngày"]]);
      return `<tr><td style="${td}">${d ? thu[d.getDay()] : ""}</td><td style="${td}">${esc_(x[C["Ngày"]])}</td>
        <td style="${td}">${esc_(caHienThi_(x[C["Ca"]]))}</td><td style="${td}">${esc_(x[C["Phòng Lab"]])}</td></tr>`;
    }).join("");
  const nguoiTao = esc_(r[C["Người tạo"]] || "Kỹ thuật viên");
  const anh = anhBanner_();

  return {
    to: r[C["Email"]],
    subject: huy ? "THÔNG BÁO HỦY LỊCH TRỰC PHÒNG LAB" : "THÔNG BÁO PHÂN CÔNG TRỰC PHÒNG LAB",
    html: thanEmail_(`
      <p>Thân gửi <b>${esc_(r[C["Họ tên"]])}</b>,</p>
      <p>${huy ? `Kỹ thuật viên ${nguoiTao} đã <b>hủy</b> lịch trực phòng Lab sau của bạn:`
               : `Kỹ thuật viên ${nguoiTao} đã phân công bạn trực phòng Lab với vai trò <b>${esc_(r[C["Vai trò"]])}</b>:`}</p>
      <table style="border-collapse:collapse;margin:8px 0 14px;font-size:14px;width:100%;max-width:620px">
        <tr style="background:#f7f7f7"><td style="${td}"><b>Thứ</b></td><td style="${td}"><b>Ngày</b></td>
          <td style="${td}"><b>Ca</b></td><td style="${td}"><b>Phòng Lab</b></td></tr>
        ${dong}
      </table>
      ${r[C["Ghi chú"]] ? `<p><b>Ghi chú:</b> ${esc_(r[C["Ghi chú"]])}</p>` : ""}
      <p>${huy ? "Bạn không cần có mặt ở các ca trên." : "Vui lòng có mặt đúng giờ. Nếu không thể tham gia, vui lòng phản hồi email này để kỹ thuật viên sắp xếp lại."}</p>`, anh),
    images: anh
  };
}

/* =====================================================================
   PHẦN DÙNG CHUNG
===================================================================== */

function moSheet_(name) {
  return SpreadsheetApp.openById(CAU_HINH.SPREADSHEET_ID).getSheetByName(name);
}

function cot_(header) {
  const C = {};
  header.forEach((h, i) => { C[String(h).trim()] = i; });
  return C;
}

function ghiO_(sheet, C) {
  return (row, col, value) => {
    if (C[col] !== undefined) sheet.getRange(row, C[col] + 1).setValue(value);
  };
}

function laCho_(r, C) {
  return C.EMAIL_STATUS !== undefined && String(r[C.EMAIL_STATUS]).trim() === EMAIL_CHO;
}

/** Gửi 1 email cho nhóm dòng rồi ghi trạng thái ĐÃ GỬI / LỖI. */
function guiMot_(sent, rows, ghi, taoEmail) {
  if (sent.n >= CAU_HINH.TOI_DA_EMAIL_MOI_LAN) return;
  const now = Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "yyyy-MM-dd HH:mm:ss");

  try {
    const mail = taoEmail();
    if (!String(mail.to || "").trim()) throw new Error("Không có email người nhận");

    GmailApp.sendEmail(mail.to, mail.subject, "", {
      htmlBody: mail.html,
      inlineImages: mail.images || {},
      attachments: mail.attachments || [],
      name: CAU_HINH.TEN_DON_VI,
      replyTo: CAU_HINH.EMAIL_DON_VI
    });
    sent.n++;
    sent.gui += rows.length;
    rows.forEach(r => { ghi(r.__row, "EMAIL_STATUS", "ĐÃ GỬI"); ghi(r.__row, "EMAIL_SENT_AT", now); ghi(r.__row, "EMAIL_ERROR", ""); });
  } catch (err) {
    sent.loi.push(String((err && err.message) || err));
    rows.forEach(r => { ghi(r.__row, "EMAIL_STATUS", "LỖI"); ghi(r.__row, "EMAIL_ERROR", String((err && err.message) || err)); });
  }
}

function bangThongTin_(dong) {
  const td = "padding:8px 12px; border:1px solid #d9d9d9;";
  return "<table style='border-collapse:collapse; margin:8px 0 16px 0; font-size:14px; width:100%; max-width:620px;'>" +
    dong.filter(x => x[1] !== "" && x[1] != null)
      .map(x => `<tr><td style="${td} background:#f7f7f7; width:42%;"><b>${x[0]}</b></td><td style="${td}">${x[1]}</td></tr>`)
      .join("") +
    "</table>";
}

/** Khung email + chữ ký giống mẫu cũ (khoa, trường, địa chỉ, banner). */
function thanEmail_(noiDung, anh) {
  const banner = anh && anh.vlu_banner
    ? "<div style='margin-top:14px;width:100%'><img src='cid:vlu_banner' alt='Van Lang University' style='display:block;width:100%;height:auto;border:0'></div>"
    : "";

  return "<div style='font-family:Arial,Helvetica,sans-serif; font-size:14px; line-height:1.7; color:#222;'>" +
    noiDung +
    "<p>Trân trọng,</p>" +
    "<div style='margin-top:20px; padding-top:14px; border-top:1px solid #d9d9d9;'>" +
      "<div style='font-size:15px; font-weight:700; color:#f0643b; text-transform:uppercase;'>KHOA KỸ THUẬT VÀ QUẢN LÝ CÔNG NGHIỆP</div>" +
      "<div style='margin-top:8px; font-size:13px;'>Phone: <a href='tel:02871098256' style='color:#1a73e8;text-decoration:none'>028.71098.256</a> (4340) | " +
        "<a href='mailto:" + CAU_HINH.EMAIL_DON_VI + "' style='color:#1a73e8;text-decoration:none'>" + CAU_HINH.EMAIL_DON_VI + "</a></div>" +
      "<div style='margin-top:6px; font-size:13px; color:#4f7df3; font-weight:700;'>TRƯỜNG ĐẠI HỌC VĂN LANG</div>" +
      "<div style='font-size:13px;'>Hotline: <a href='tel:02871059999' style='color:#1a73e8;text-decoration:none'>028.7105.9999</a> | " +
        "<a href='mailto:info@vlu.edu.vn' style='color:#1a73e8;text-decoration:none'>info@vlu.edu.vn</a> | " +
        "<a href='https://vlu.edu.vn' style='color:#1a73e8;text-decoration:none'>vlu.edu.vn</a></div>" +
      "<div style='margin-top:8px; font-size:13px; line-height:1.7;'>" +
        "<b>CƠ SỞ CHÍNH:</b> 69/68 Đặng Thùy Trâm, Phường Bình Lợi Trung, TP. Hồ Chí Minh<br>" +
        "<b>CS1:</b> 45 Nguyễn Khắc Nhu, Phường Cầu Ông Lãnh, TP. Hồ Chí Minh<br>" +
        "<b>CS2:</b> 233A Phan Văn Trị, Phường Bình Lợi Trung, TP. Hồ Chí Minh</div>" +
      banner +
    "</div></div>";
}

function anhBanner_() {
  try {
    return CAU_HINH.BANNER_FILE_ID ? { vlu_banner: DriveApp.getFileById(CAU_HINH.BANNER_FILE_ID).getBlob() } : {};
  } catch (err) {
    Logger.log("Không lấy được banner email (kiểm tra quyền xem file của tài khoản chạy script): " + err);
    return {};
  }
}

/** File lịch .ics đính kèm (mỗi ca 1 sự kiện) – giống mẫu cũ. */
function fileLich_(rows, C) {
  const fmt = d => Utilities.formatDate(d, "UTC", "yyyyMMdd'T'HHmmss'Z'");
  const icsText = t => String(t || "").replace(/\\/g, "\\\\").replace(/\n/g, "\\n").replace(/,/g, "\\,").replace(/;/g, "\\;");
  const events = rows.map(r => {
    const d = ngayDate_(r[C.NGAY_SU_DUNG]);
    const gio = CAU_HINH.GIO_CA[String(r[C.CA_SU_DUNG]).trim()];
    if (!d || !gio) return "";
    const at = hm => new Date(d.getFullYear(), d.getMonth(), d.getDate(), Number(hm.split(":")[0]), Number(hm.split(":")[1]));
    return "BEGIN:VEVENT\r\n" +
      "UID:lab-" + r[C.ID_KEY] + "-" + String(r[C.CA_SU_DUNG]).replace(/\s+/g, "") + "@vlu.edu.vn\r\n" +
      "DTSTAMP:" + fmt(new Date()) + "\r\n" +
      "DTSTART:" + fmt(at(gio[0])) + "\r\nDTEND:" + fmt(at(gio[1])) + "\r\n" +
      "SUMMARY:" + icsText(r[C.PHONG] + " | " + r[C.CA_SU_DUNG] + " | " + r[C.ID_KEY]) + "\r\n" +
      "LOCATION:" + icsText(CAU_HINH.DIA_DIEM) + "\r\n" +
      "DESCRIPTION:" + icsText("ID_KEY: " + r[C.ID_KEY] + "\nKhu vực: " + r[C.KHU_VUC]) + "\r\n" +
      "END:VEVENT\r\n";
  }).join("");

  return Utilities.newBlob(
    "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//Van Lang University//Lab Booking//VI\r\nCALSCALE:GREGORIAN\r\nMETHOD:PUBLISH\r\n" +
    events + "END:VCALENDAR", "text/calendar", "lich-dang-ky-lab.ics");
}

function caHienThi_(ca) {
  const ten = String(ca || "").trim();
  const gio = CAU_HINH.GIO_CA[ten];
  return gio ? `${ten} (${gio[0]} - ${gio[1]})` : ten;
}

function chuanSlot_(v) {
  return String(v || "").trim().replace(/\s+/g, " ").toLowerCase();
}

/** Ngày dạng 2026-10-30 hoặc 30/10/2026 -> Date */
function ngayDate_(v) {
  const s = String(v || "").trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  return null;
}

function ngayVN_(v) {
  const d = ngayDate_(v);
  return d ? Utilities.formatDate(d, "Asia/Ho_Chi_Minh", "dd/MM/yyyy") : String(v || "");
}

function esc_(text) {
  return String(text == null ? "" : text)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#039;");
}

/* =====================================================================
   CÀI ĐẶT & KIỂM TRA
===================================================================== */

/** Chạy 1 lần: xóa trình kích hoạt cũ (kể cả của Google Form), tạo 1 trình kích hoạt guiEmail mỗi 1 phút,
 *  rồi gửi luôn các email đang chờ. */
function caiDatTuDong() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("guiEmail").timeBased().everyMinutes(1).create();
  kiemTraHeThong();
  guiEmail();
}

/** Ghi ra Nhật ký thực thi: sheet, số email đang chờ, trình kích hoạt, hạn mức gửi còn lại. */
function kiemTraHeThong() {
  [CAU_HINH.SHEET_DANG_KY, CAU_HINH.SHEET_LICH_TRUC].forEach(name => {
    const sheet = moSheet_(name);
    if (!sheet) { Logger.log("THIẾU sheet " + name); return; }
    const values = sheet.getDataRange().getDisplayValues();
    const C = cot_(values[0]);
    if (C.EMAIL_STATUS === undefined) { Logger.log(name + ": THIẾU cột EMAIL_STATUS"); return; }
    const dem = {};
    values.slice(1).forEach(r => { const k = String(r[C.EMAIL_STATUS]).trim() || "(trống)"; dem[k] = (dem[k] || 0) + 1; });
    Logger.log(name + ": " + JSON.stringify(dem));
  });
  Logger.log("Trình kích hoạt: " + ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction()).join(", "));
  Logger.log("Số email còn gửi được hôm nay: " + MailApp.getRemainingDailyQuota());
  Logger.log("Banner: " + (anhBanner_().vlu_banner ? "OK" : "không lấy được"));
}
