"""
MƯỢN – TRẢ VẬT TƯ & NHẬT KÝ SỰ CỐ – BẢO TRÌ (Dashboard kỹ thuật viên)
=====================================================================
Theo SOP-03 / SOP-04 bản 2.0 (Lab 1):
- Món trên 1 triệu: sinh viên báo miệng, KTV ghi mượn – trả tại trang "Mượn – trả".
  Sinh viên không phải điền phiếu. Món dưới 1 triệu tự lấy, không cần ghi.
- Nhật ký sự cố: tự thêm dòng khi
    + phiếu bàn giao chọn "Có sự cố cần xử lý" / có phát sinh bất thường,
    + KTV nhận trả món bị Hỏng / Thiếu.
  KTV bổ sung mã máy, cách xử lý, ngày xong, trạng thái. KTV tự phát hiện: bấm "Thêm sự cố".
- Vật tư (Nội dung 4 của thuyết minh: theo dõi thiết bị, vật tư):
    + Danh mục món theo mã vị trí (T1, T2, K1 – K3, D4), nhóm Tự lấy / Báo KTV,
      SL chuẩn, SL tối thiểu. KTV nhập hoặc dán từ Excel; hệ thống không tự điền số lượng.
    + Kiểm kê: mỗi lần KTV đếm thực tế ghi 1 dòng (SL thực tế, số hỏng).
    + Tình trạng tự tính: đang cho mượn (từ MUON_TRA), hỏng chờ xử lý (từ NHAT_KY_SU_CO),
      dưới mức tối thiểu, thiếu so với chuẩn, quá hạn kiểm kê.
- Mượn – trả đối chiếu phiếu bàn giao: lượt đã bàn giao phòng mà còn món chưa trả -> cảnh báo.
- Dữ liệu lưu trong file Google Sheet Booking (cùng file với LAB_ATTENDANCE),
  4 sheet tự tạo: MUON_TRA, NHAT_KY_SU_CO, VAT_TU, KIEM_KE.

Cài vào app.py (thêm 2 dòng, ngay trước phần "RUN"):
    from lab_ops import register_lab_ops
    register_lab_ops(app, globals())
"""

import re
import threading
import uuid
from datetime import datetime, timedelta

from flask import jsonify, redirect, render_template, request, session, url_for

LOAN_SHEET = "MUON_TRA"
LOAN_HEADERS = [
    "Mã mượn", "Thời gian mượn", "Mã đăng ký", "MSSV", "Họ tên", "Món / thiết bị",
    "Mã vị trí", "Số lượng", "Tình trạng khi mượn", "KTV giao", "Trạng thái",
    "Thời gian trả", "Tình trạng khi trả", "KTV nhận", "Ghi chú", "Mã vật tư",
]
LOAN_OPEN = "ĐANG MƯỢN"
LOAN_DONE = "ĐÃ TRẢ"
RETURN_CONDITIONS = ["Bình thường", "Hỏng", "Thiếu"]

INCIDENT_SHEET = "NHAT_KY_SU_CO"
INCIDENT_HEADERS = [
    "Mã sự cố", "Ngày báo", "Phòng", "Mã máy / vị trí", "Loại sự cố", "Mô tả",
    "Người báo", "Mã đăng ký", "Nguồn", "Trạng thái", "Cách xử lý", "Ngày xong",
    "KTV xử lý", "Cập nhật lúc",
]
INCIDENT_STATUSES = ["Mới", "Đang xử lý", "Đã xong"]
INCIDENT_TYPES = ["Hỏng máy / thiết bị", "Thiếu, mất vật tư", "Mất vệ sinh, sai vị trí", "An toàn (cháy, điện, chấn thương)", "Khác"]
SOURCE_HANDOVER = "Phiếu bàn giao"
SOURCE_RETURN = "Trả đồ"
SOURCE_STAFF = "KTV ghi"

# Phiếu bàn giao cũ hơn số ngày này không tự đưa vào nhật ký (tránh dồn dữ liệu cũ)
AUTO_SYNC_DAYS = 120

# Món trên 1 triệu của Lab 1 (gợi ý khi nhập; KTV vẫn gõ được món khác)
VALUABLE_ITEMS = [
    {"name": "Máy mài Bosch", "code": "K1-N2-06"},
    {"name": "Máy cắt rung Bosch", "code": "K1-N2-06"},
    {"name": "Máy khoan pin Bosch", "code": "T1-N3"},
    {"name": "Pin Bosch 18V", "code": "T1-N3"},
    {"name": "Sạc pin Bosch", "code": "T1-N2"},
    {"name": "Máy hàn", "code": "T1-N2"},
    {"name": "Đồng hồ vạn năng (VOM)", "code": "T1-N2"},
    {"name": "Hộp máy móc – kệ K1", "code": "K1-N1"},
    {"name": "Hộp máy móc – kệ K2", "code": "K2-N1"},
]

# Máy khu C của Lab 1 (theo sơ đồ 5S, decal máy và trang Hướng dẫn của cổng sinh viên)
MACHINES = [
    {"code": "C1", "name": "Máy in 3D"},
    {"code": "C2", "name": "Phay CNC mini"},
    {"code": "C3", "name": "Tiện mini"},
    {"code": "C4", "name": "Cưa lọng bàn"},
    {"code": "C5", "name": "Máy cắt góc"},
    {"code": "C6", "name": "Khoan bàn"},
    {"code": "C7", "name": "Máy cắt"},
]

INV_SHEET = "VAT_TU"
INV_HEADERS = [
    "Mã vật tư", "Tên vật tư", "Nhóm", "Mã vị trí", "Đơn vị", "SL chuẩn", "SL tối thiểu",
    "Ghi chú", "Trạng thái", "Cập nhật lúc", "Người cập nhật",
]
INV_GROUPS = ["Báo KTV", "Tự lấy"]          # Báo KTV: trên 1 triệu / món (SOP-03)
INV_ACTIVE = "Đang dùng"
INV_RETIRED = "Ngừng dùng"
COUNT_SHEET = "KIEM_KE"
COUNT_HEADERS = [
    "Mã kiểm kê", "Thời gian", "Mã vật tư", "Tên vật tư", "Mã vị trí", "SL thực tế", "Số hỏng",
    "Ghi chú", "KTV",
]
STORAGE_LOCATIONS = ["T1", "T2", "K1", "K2", "K3", "D4"]
COUNT_DUE_DAYS = 31                          # SOP-03: kiểm kê mỗi tháng

LAB_POSITIONS = ["C1", "C2", "C3", "C4", "C5", "C6", "C7", "B1", "B2", "B3", "B4", "B5",
                 "B6", "B7", "B8", "B9", "B10", "T1", "T2", "K1", "K2", "K3", "F1", "F2"]

_lock = threading.Lock()


def register_lab_ops(app, h):
    """h = globals() của app.py (dùng lại kết nối Google Sheet, bộ nhớ đệm, đăng nhập)."""

    clean = h["clean_value"]
    get_sheet = h["get_or_create_sheet"]
    invalidate = h["invalidate_cache"]
    values_cached = h["sheet_values_cached"]
    require_admin = h["require_admin_api"]
    pick = h["get_value_by_possible_keys"]
    no_accent = h["remove_vietnamese_accents"]

    def now_text():
        return datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    def parse_dt(text):
        text = clean(text)
        for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
            try:
                return datetime.strptime(text, fmt)
            except ValueError:
                pass
        return None

    def read(sheet_name, headers, fresh=False):
        sheet = get_sheet(sheet_name, headers)
        values = sheet.get_all_values() if fresh else values_cached(sheet, sheet_name)
        if values and [clean(v) for v in values[0][:len(headers)]] != headers:
            sheet.update(range_name="A1", values=[headers], value_input_option="RAW")
        rows = []
        for index, raw in enumerate(values[1:], start=2):
            raw = list(raw) + [""] * (len(headers) - len(raw))
            if not clean(raw[0]):
                continue
            item = {header: clean(raw[i]) for i, header in enumerate(headers)}
            item["_row"] = index
            rows.append(item)
        return sheet, rows

    def write_row(sheet, row_number, headers, item):
        sheet.update(range_name=f"A{row_number}", values=[[item.get(h_, "") for h_ in headers]],
                     value_input_option="RAW")

    def staff_name():
        return clean(session.get("user")) or h.get("TECH_NAME", "KTV")

    def admin_page(template, **extra):
        if "user" not in session:
            return redirect(url_for("login"))
        if session.get("role") == "lecturer":
            return redirect(url_for("teacher"))
        return render_template(template, user=session["user"], **extra)

    # =====================================================
    # TRANG
    # =====================================================
    @app.route("/loans")
    def loans_page():
        return admin_page("loans.html", valuable_items=VALUABLE_ITEMS)

    @app.route("/incidents")
    def incidents_page():
        return admin_page("incidents.html", incident_types=INCIDENT_TYPES,
                          incident_statuses=INCIDENT_STATUSES, positions=LAB_POSITIONS, machines=MACHINES)

    @app.route("/inventory")
    def inventory_page():
        return admin_page("inventory.html", groups=INV_GROUPS, locations=STORAGE_LOCATIONS,
                          seed_items=VALUABLE_ITEMS)

    # =====================================================
    # HÀM DÙNG CHUNG
    # =====================================================
    def norm_name(text):
        return " ".join(no_accent(clean(text)).lower().split())

    def to_int(value):
        text = clean(value).replace(",", ".")
        if not text:
            return None
        try:
            return max(0, int(float(text)))
        except ValueError:
            return None

    def location_root(code):
        match = re.match(r"^(T[12]|K[1-3]|D[1-4])", clean(code).upper())
        return match.group(1) if match else ""

    def handed_over_ids():
        """Mã đăng ký đã có phiếu bàn giao -> thời gian bàn giao (form cũ + cổng sinh viên)."""
        found = {}
        try:
            for row in h["get_lab_handover_rows"]():
                id_key = pick(row, ["1. ID_KEY của bạn là:", "ID_KEY", "ID_KEY của bạn là:"]).upper()
                if id_key:
                    found[id_key] = pick(row, ["Dấu thời gian", "UPDATED_AT"])
        except Exception as e:
            print("HANDOVER READ ERROR:", e)
        return found

    # =====================================================
    # MƯỢN – TRẢ
    # =====================================================
    def loan_public(item, handed=None):
        borrowed = parse_dt(item["Thời gian mượn"])
        is_open = item["Trạng thái"] == LOAN_OPEN
        overdue = is_open and borrowed is not None and borrowed.date() < datetime.now().date()
        handed_at = (handed or {}).get(item["Mã đăng ký"].upper(), None) if item["Mã đăng ký"] else None
        return {
            "item_id": item["Mã vật tư"],
            "handed_over": bool(is_open and handed_at is not None),
            "handed_at": handed_at or "",
            "id": item["Mã mượn"],
            "borrowed_at": item["Thời gian mượn"],
            "booking_id": item["Mã đăng ký"],
            "student_id": item["MSSV"],
            "name": item["Họ tên"],
            "item": item["Món / thiết bị"],
            "location": item["Mã vị trí"],
            "quantity": item["Số lượng"],
            "condition_out": item["Tình trạng khi mượn"],
            "staff_out": item["KTV giao"],
            "status": item["Trạng thái"],
            "returned_at": item["Thời gian trả"],
            "condition_in": item["Tình trạng khi trả"],
            "staff_in": item["KTV nhận"],
            "note": item["Ghi chú"],
            "overdue": overdue,
        }

    @app.route("/api/loans")
    def api_loans():
        denied = require_admin()
        if denied:
            return denied
        try:
            _, rows = read(LOAN_SHEET, LOAN_HEADERS)
            handed = handed_over_ids() if any(r["Trạng thái"] == LOAN_OPEN for r in rows) else {}
            data = [loan_public(item, handed) for item in rows]
            data.sort(key=lambda x: (x["status"] != LOAN_OPEN, -(parse_dt(x["borrowed_at"]) or datetime.min).timestamp()))
            catalog = [
                {"id": x["id"], "name": x["name"], "code": x["location"], "unit": x["unit"],
                 "available": x["available"], "borrowed": x["borrowed"]}
                for x in inventory_view(loans=rows) if x["group"] == "Báo KTV" and x["active"]
            ]
            if not catalog:
                catalog = [{"id": "", "name": x["name"], "code": x["code"], "unit": "", "available": None, "borrowed": 0}
                           for x in VALUABLE_ITEMS]
            return jsonify({"success": True, "data": data, "items": catalog})
        except Exception as e:
            print("API LOANS ERROR:", e)
            return jsonify({"success": False, "data": [], "error": str(e)}), 500

    @app.route("/api/loans/lookup")
    def api_loans_lookup():
        """Mã đăng ký -> danh sách sinh viên của lượt (tự điền họ tên, MSSV)."""
        denied = require_admin()
        if denied:
            return denied
        id_key = clean(request.args.get("id_key")).upper()
        if not id_key:
            return jsonify({"success": False, "error": "Nhập mã đăng ký."}), 400
        try:
            people, seen, info = [], set(), {}
            for record in h["build_attendance_records"]():
                if clean(record.get("booking_id")).upper() != id_key:
                    continue
                info = {"lab": record.get("lab", ""), "date": record.get("date", ""), "shift": record.get("shift", "")}
                key = clean(record.get("student_id")) or clean(record.get("name"))
                if key and key not in seen:
                    seen.add(key)
                    people.append({"student_id": clean(record.get("student_id")), "name": clean(record.get("name")),
                                   "role": record.get("role", "")})
            if not people:
                return jsonify({"success": False, "error": "Không tìm thấy lượt đăng ký còn hiệu lực với mã này."}), 404
            return jsonify({"success": True, "people": people, **info})
        except Exception as e:
            print("API LOAN LOOKUP ERROR:", e)
            return jsonify({"success": False, "error": str(e)}), 500

    @app.route("/api/loans/sessions")
    def api_loans_sessions():
        """Các lượt đăng ký của hôm nay (để KTV bấm chọn khi ghi mượn)."""
        denied = require_admin()
        if denied:
            return denied
        try:
            today = datetime.now().date()
            sessions = {}
            for record in h["build_attendance_records"]():
                day = parse_dt(record.get("date"))
                booking_id = clean(record.get("booking_id")).upper()
                if not booking_id or not day or day.date() != today:
                    continue
                item = sessions.setdefault(booking_id, {
                    "booking_id": booking_id, "lab": record.get("lab", ""), "date": record.get("date", ""),
                    "shifts": [], "people": [], "_seen": set()})
                shift = clean(record.get("shift"))
                if shift and shift not in item["shifts"]:
                    item["shifts"].append(shift)
                key = clean(record.get("student_id")) or clean(record.get("name"))
                if key and key not in item["_seen"]:
                    item["_seen"].add(key)
                    item["people"].append({"student_id": clean(record.get("student_id")), "name": clean(record.get("name")),
                                           "role": record.get("role", ""),
                                           "attendance": record.get("attendance_status", "")})
            data = []
            for item in sessions.values():
                item.pop("_seen")
                item["shifts"].sort()
                item["shift"] = ", ".join(item["shifts"])
                data.append(item)
            data.sort(key=lambda x: (x["shift"], x["booking_id"]))
            return jsonify({"success": True, "data": data})
        except Exception as e:
            print("API LOAN SESSIONS ERROR:", e)
            return jsonify({"success": False, "data": [], "error": str(e)}), 500

    @app.route("/api/loans", methods=["POST"])
    def api_loans_add():
        denied = require_admin()
        if denied:
            return denied
        p = request.get_json(silent=True) or {}
        item_name = " ".join(clean(p.get("item")).split())
        name = " ".join(clean(p.get("name")).split())
        student_id = re.sub(r"\D", "", clean(p.get("student_id")))
        try:
            quantity = max(1, int(float(clean(p.get("quantity")) or 1)))
        except ValueError:
            quantity = 1
        item_id = clean(p.get("item_id")).upper()
        if item_id:
            known = next((x for x in inventory_view() if x["id"] == item_id), None)
            if known:
                item_name = item_name or known["name"]
                if not clean(p.get("location")):
                    p["location"] = known["location"]
        if not item_name:
            return jsonify({"success": False, "error": "Nhập tên món / thiết bị."}), 400
        if not name and not student_id:
            return jsonify({"success": False, "error": "Nhập họ tên hoặc MSSV người mượn."}), 400
        if student_id and len(student_id) != 13:
            return jsonify({"success": False, "error": "MSSV phải gồm đúng 13 chữ số."}), 400
        try:
            with _lock:
                sheet = get_sheet(LOAN_SHEET, LOAN_HEADERS)
                loan_id = "MT-" + datetime.now().strftime("%y%m%d") + "-" + uuid.uuid4().hex[:4].upper()
                row = {
                    "Mã mượn": loan_id, "Thời gian mượn": now_text(),
                    "Mã đăng ký": clean(p.get("booking_id")).upper(), "MSSV": student_id, "Họ tên": name,
                    "Món / thiết bị": item_name, "Mã vị trí": clean(p.get("location")).upper(),
                    "Số lượng": quantity, "Tình trạng khi mượn": clean(p.get("condition_out")) or "Bình thường",
                    "KTV giao": staff_name(), "Trạng thái": LOAN_OPEN, "Ghi chú": clean(p.get("note")),
                    "Mã vật tư": item_id,
                }
                sheet.append_row([row.get(h_, "") for h_ in LOAN_HEADERS], value_input_option="RAW")
                invalidate(LOAN_SHEET)
            return jsonify({"success": True, "id": loan_id, "message": f"Đã ghi mượn {item_name} ({loan_id})."})
        except Exception as e:
            print("API LOAN ADD ERROR:", e)
            return jsonify({"success": False, "error": str(e)}), 500

    @app.route("/api/loans/<loan_id>/return", methods=["POST"])
    def api_loans_return(loan_id):
        denied = require_admin()
        if denied:
            return denied
        p = request.get_json(silent=True) or {}
        condition = clean(p.get("condition")) or "Bình thường"
        note = clean(p.get("note"))
        if condition not in RETURN_CONDITIONS:
            return jsonify({"success": False, "error": "Tình trạng khi trả không hợp lệ."}), 400
        if condition != "Bình thường" and not note:
            return jsonify({"success": False, "error": "Món bị hỏng / thiếu: ghi mô tả để đưa vào nhật ký sự cố."}), 400
        try:
            with _lock:
                sheet, rows = read(LOAN_SHEET, LOAN_HEADERS, fresh=True)
                item = next((r for r in rows if r["Mã mượn"] == clean(loan_id)), None)
                if not item:
                    return jsonify({"success": False, "error": "Không tìm thấy lượt mượn."}), 404
                if item["Trạng thái"] == LOAN_DONE:
                    return jsonify({"success": False, "error": "Lượt này đã ghi trả."}), 409
                item.update({"Trạng thái": LOAN_DONE, "Thời gian trả": now_text(), "Tình trạng khi trả": condition,
                             "KTV nhận": staff_name(),
                             "Ghi chú": "; ".join(x for x in [item["Ghi chú"], note] if x)})
                write_row(sheet, item["_row"], LOAN_HEADERS, item)
                invalidate(LOAN_SHEET)

                incident_id = ""
                if condition != "Bình thường":
                    isheet = get_sheet(INCIDENT_SHEET, INCIDENT_HEADERS)
                    incident_id = "TD-" + item["Mã mượn"]
                    isheet.append_row([{
                        "Mã sự cố": incident_id, "Ngày báo": now_text(), "Phòng": "Lab 1",
                        "Mã máy / vị trí": item["Mã vị trí"],
                        "Loại sự cố": "Hỏng máy / thiết bị" if condition == "Hỏng" else "Thiếu, mất vật tư",
                        "Mô tả": f"{item['Món / thiết bị']} trả về {condition.lower()}: {note}",
                        "Người báo": " · ".join(x for x in [item["Họ tên"], item["MSSV"]] if x),
                        "Mã đăng ký": item["Mã đăng ký"], "Nguồn": SOURCE_RETURN, "Trạng thái": "Mới",
                        "Cập nhật lúc": now_text(),
                    }.get(h_, "") for h_ in INCIDENT_HEADERS], value_input_option="RAW")
                    invalidate(INCIDENT_SHEET)
            message = "Đã ghi trả."
            if incident_id:
                message += f" Món {condition.lower()} đã chuyển sang Nhật ký sự cố ({incident_id})."
            return jsonify({"success": True, "message": message, "incident_id": incident_id})
        except Exception as e:
            print("API LOAN RETURN ERROR:", e)
            return jsonify({"success": False, "error": str(e)}), 500

    # =====================================================
    # NHẬT KÝ SỰ CỐ – BẢO TRÌ
    # =====================================================
    def handover_problems():
        """Phiếu bàn giao có sự cố / bất thường (Google Form cũ + cổng sinh viên)."""
        found = {}
        limit = datetime.now() - timedelta(days=AUTO_SYNC_DAYS)
        for row in h["get_lab_handover_rows"]():
            id_key = pick(row, ["1. ID_KEY của bạn là:", "ID_KEY", "ID_KEY của bạn là:"])
            if not id_key:
                continue
            abnormal = pick(row, [h.get("FORM_ABNORMAL_KEY", ""),
                                  "Phòng Lab có phát sinh vấn đề bất thường trong quá trình sử dụng không?"])
            condition = pick(row, ["TINH_TRANG"])
            error = pick(row, ["ERROR_MESSAGE"])
            text = no_accent(f"{abnormal} {condition}").lower()
            if not (error or "su co" in text or no_accent(abnormal).lower().startswith("co")):
                continue
            when = parse_dt(pick(row, ["Dấu thời gian", "UPDATED_AT"]))
            if when and when < limit:
                continue
            code = pick(row, ["MA_BAN_GIAO"]) or id_key
            batch = pick(row, ["BAN_GIAO_CHUNG_VOI"]) or id_key
            key = f"BG-{code}"
            if key in found:
                continue
            description = error or " · ".join(x for x in [condition, pick(row, ["GHI_CHU"])] if x) or "Phiếu bàn giao báo có bất thường"
            # Đoán mã máy / vị trí từ mô tả của sinh viên (vd "máy khoan C6 kêu to" -> C6); KTV sửa lại được
            codes = re.findall(r"\b(C[1-7]|B(?:10|[1-9])|T[12]|K[1-3]|F[12])\b", description.upper())
            found[key] = {
                "Mã sự cố": key,
                "Ngày báo": when.strftime("%d/%m/%Y %H:%M:%S") if when else now_text(),
                "Phòng": pick(row, ["7. Phòng Lab bạn đã sử dụng là:", "9. Phòng Lab bạn đã sử dụng là:", "Phòng", "Lab"])
                         or h["extract_lab_from_id_key"](id_key),
                "Mã máy / vị trí": ", ".join(dict.fromkeys(codes)),
                "Loại sự cố": "Khác",
                "Mô tả": description,
                "Người báo": " · ".join(x for x in [pick(row, ["3. Họ và tên của bạn là:", "5. Họ và tên của bạn là:", "Họ và tên"]),
                                                     pick(row, ["MSSV"])] if x),
                "Mã đăng ký": batch,
                "Nguồn": SOURCE_HANDOVER, "Trạng thái": "Mới", "Cập nhật lúc": now_text(),
            }
        return found

    def sync_from_handover():
        """Thêm vào sheet các phiếu bàn giao có sự cố chưa có trong nhật ký."""
        try:
            problems = handover_problems()
        except Exception as e:
            print("INCIDENT SYNC READ ERROR:", e)
            return 0
        if not problems:
            return 0
        with _lock:
            sheet, rows = read(INCIDENT_SHEET, INCIDENT_HEADERS, fresh=True)
            have = {r["Mã sự cố"] for r in rows}
            new = [v for k, v in problems.items() if k not in have]
            if new:
                sheet.append_rows([[item.get(h_, "") for h_ in INCIDENT_HEADERS] for item in new],
                                  value_input_option="RAW")
                invalidate(INCIDENT_SHEET)
        return len(new)

    def incident_public(item):
        reported = parse_dt(item["Ngày báo"])
        done = parse_dt(item["Ngày xong"])
        hours = round((done - reported).total_seconds() / 3600, 1) if (reported and done and done >= reported) else None
        return {
            "id": item["Mã sự cố"], "reported_at": item["Ngày báo"], "lab": item["Phòng"],
            "position": item["Mã máy / vị trí"], "type": item["Loại sự cố"], "description": item["Mô tả"],
            "reporter": item["Người báo"], "booking_id": item["Mã đăng ký"], "source": item["Nguồn"],
            "status": item["Trạng thái"] or "Mới", "action": item["Cách xử lý"], "done_at": item["Ngày xong"],
            "staff": item["KTV xử lý"], "updated_at": item["Cập nhật lúc"], "hours_to_fix": hours,
        }

    @app.route("/api/incidents")
    def api_incidents():
        denied = require_admin()
        if denied:
            return denied
        try:
            added = sync_from_handover() if request.args.get("sync", "1") != "0" else 0
            _, rows = read(INCIDENT_SHEET, INCIDENT_HEADERS)
            data = [incident_public(r) for r in rows]
            order = {s: i for i, s in enumerate(INCIDENT_STATUSES)}
            data.sort(key=lambda x: (order.get(x["status"], 0), -(parse_dt(x["reported_at"]) or datetime.min).timestamp()))
            return jsonify({"success": True, "data": data, "synced": added})
        except Exception as e:
            print("API INCIDENTS ERROR:", e)
            return jsonify({"success": False, "data": [], "error": str(e)}), 500

    @app.route("/api/incidents", methods=["POST"])
    def api_incidents_add():
        denied = require_admin()
        if denied:
            return denied
        p = request.get_json(silent=True) or {}
        description = clean(p.get("description"))
        if not description:
            return jsonify({"success": False, "error": "Nhập mô tả sự cố."}), 400
        try:
            with _lock:
                sheet = get_sheet(INCIDENT_SHEET, INCIDENT_HEADERS)
                incident_id = "SC-" + datetime.now().strftime("%y%m%d") + "-" + uuid.uuid4().hex[:4].upper()
                item = {
                    "Mã sự cố": incident_id, "Ngày báo": now_text(), "Phòng": clean(p.get("lab")) or "Lab 1",
                    "Mã máy / vị trí": clean(p.get("position")).upper(),
                    "Loại sự cố": clean(p.get("type")) or "Khác", "Mô tả": description,
                    "Người báo": clean(p.get("reporter")) or staff_name(), "Mã đăng ký": clean(p.get("booking_id")).upper(),
                    "Nguồn": SOURCE_STAFF, "Trạng thái": "Mới", "Cập nhật lúc": now_text(),
                }
                sheet.append_row([item.get(h_, "") for h_ in INCIDENT_HEADERS], value_input_option="RAW")
                invalidate(INCIDENT_SHEET)
            return jsonify({"success": True, "id": incident_id, "message": f"Đã thêm sự cố {incident_id}."})
        except Exception as e:
            print("API INCIDENT ADD ERROR:", e)
            return jsonify({"success": False, "error": str(e)}), 500

    @app.route("/api/incidents/<incident_id>", methods=["POST"])
    def api_incidents_update(incident_id):
        denied = require_admin()
        if denied:
            return denied
        p = request.get_json(silent=True) or {}
        status = clean(p.get("status"))
        if status and status not in INCIDENT_STATUSES:
            return jsonify({"success": False, "error": "Trạng thái không hợp lệ."}), 400
        try:
            with _lock:
                sheet, rows = read(INCIDENT_SHEET, INCIDENT_HEADERS, fresh=True)
                item = next((r for r in rows if r["Mã sự cố"] == clean(incident_id)), None)
                if not item:
                    return jsonify({"success": False, "error": "Không tìm thấy sự cố."}), 404
                for field, key in (("position", "Mã máy / vị trí"), ("type", "Loại sự cố"), ("action", "Cách xử lý"),
                                   ("description", "Mô tả")):
                    if field in p:
                        value = clean(p.get(field))
                        item[key] = value.upper() if field == "position" else value
                if status:
                    item["Trạng thái"] = status
                if status == "Đã xong":
                    if not item["Cách xử lý"]:
                        return jsonify({"success": False, "error": "Ghi cách xử lý trước khi chuyển Đã xong."}), 400
                    item["Ngày xong"] = item["Ngày xong"] or now_text()
                elif status:
                    item["Ngày xong"] = ""
                item["KTV xử lý"] = staff_name()
                item["Cập nhật lúc"] = now_text()
                write_row(sheet, item["_row"], INCIDENT_HEADERS, item)
                invalidate(INCIDENT_SHEET)
            return jsonify({"success": True, "message": f"Đã cập nhật {incident_id}."})
        except Exception as e:
            print("API INCIDENT UPDATE ERROR:", e)
            return jsonify({"success": False, "error": str(e)}), 500

    # =====================================================
    # VẬT TƯ – KIỂM KÊ
    # =====================================================
    def inventory_view(loans=None):
        """Danh mục + tình trạng tự tính từ kiểm kê, mượn – trả và nhật ký sự cố."""
        _, items = read(INV_SHEET, INV_HEADERS)
        if not items:
            return []
        if loans is None:
            _, loans = read(LOAN_SHEET, LOAN_HEADERS)
        _, counts = read(COUNT_SHEET, COUNT_HEADERS)
        _, incidents = read(INCIDENT_SHEET, INCIDENT_HEADERS)

        by_id = {x["Mã vật tư"].upper(): x for x in items}
        by_name = {}
        for x in items:
            by_name.setdefault(norm_name(x["Tên vật tư"]), x["Mã vật tư"].upper())

        def item_of_loan(loan):
            code = loan["Mã vật tư"].upper()
            return code if code in by_id else by_name.get(norm_name(loan["Món / thiết bị"]), "")

        borrowed, loan_item = {}, {}
        for loan in loans:
            code = item_of_loan(loan)
            loan_item[loan["Mã mượn"]] = code
            if code and loan["Trạng thái"] == LOAN_OPEN:
                borrowed[code] = borrowed.get(code, 0) + (to_int(loan["Số lượng"]) or 1)

        broken_open = {}
        for inc in incidents:
            if inc["Trạng thái"] == "Đã xong" or not inc["Mã sự cố"].startswith("TD-"):
                continue
            code = loan_item.get(inc["Mã sự cố"][3:], "")
            if code and "hỏng" in inc["Mô tả"].lower():
                broken_open[code] = broken_open.get(code, 0) + 1

        last_count = {}
        for c in counts:
            code = c["Mã vật tư"].upper()
            when = parse_dt(c["Thời gian"]) or datetime.min
            if code not in last_count or when >= last_count[code][0]:
                last_count[code] = (when, c)

        now = datetime.now()
        data = []
        for x in items:
            code = x["Mã vật tư"].upper()
            std, minimum = to_int(x["SL chuẩn"]), to_int(x["SL tối thiểu"])
            out = borrowed.get(code, 0)
            last = last_count.get(code)
            qty = to_int(last[1]["SL thực tế"]) if last else None
            broken = (to_int(last[1]["Số hỏng"]) or 0) if last else 0
            broken += broken_open.get(code, 0)
            base = qty if qty is not None else std
            available = max(0, base - out - broken) if base is not None else None
            active = x["Trạng thái"] != INV_RETIRED
            days = (now - last[0]).days if last and last[0] != datetime.min else None

            if not active:
                status, tone = "Ngừng dùng", "gray"
            elif last is None:
                status, tone = "Chưa kiểm kê", "gray"
            elif minimum is not None and qty is not None and qty <= minimum:
                status, tone = "Dưới mức tối thiểu", "red"
            elif std is not None and qty is not None and qty < std:
                status, tone = "Thiếu so với chuẩn", "amber"
            elif broken:
                status, tone = "Có món hỏng", "amber"
            elif days is not None and days > COUNT_DUE_DAYS:
                status, tone = "Cần kiểm kê lại", "amber"
            else:
                status, tone = "Đủ", "green"

            data.append({
                "id": code, "name": x["Tên vật tư"], "group": x["Nhóm"] or "Tự lấy",
                "location": x["Mã vị trí"], "root": location_root(x["Mã vị trí"]), "unit": x["Đơn vị"],
                "std": std, "min": minimum, "note": x["Ghi chú"], "active": active,
                "borrowed": out, "broken": broken, "available": available,
                "last_qty": qty, "last_count_at": last[1]["Thời gian"] if last else "",
                "last_count_by": last[1]["KTV"] if last else "", "days_since_count": days,
                "status": status, "tone": tone, "updated_at": x["Cập nhật lúc"],
            })
        data.sort(key=lambda d: (not d["active"], STORAGE_LOCATIONS.index(d["root"]) if d["root"] in STORAGE_LOCATIONS else 99,
                                 d["location"], d["name"]))
        return data

    def next_item_code(items):
        numbers = [int(m.group(1)) for x in items for m in [re.match(r"^VT-(\d+)$", x["Mã vật tư"].upper())] if m]
        return f"VT-{(max(numbers) if numbers else 0) + 1:03d}"

    def item_row(p, old=None):
        """Kiểm tra dữ liệu 1 món, trả về (dòng, lỗi)."""
        old = old or {}
        name = " ".join(clean(p.get("name", old.get("Tên vật tư", ""))).split())
        group = clean(p.get("group", old.get("Nhóm", ""))) or "Tự lấy"
        if no_accent(group).lower().replace(" ", "") in ("baoktv", "nhom2", "2"):
            group = "Báo KTV"
        elif no_accent(group).lower().replace(" ", "") in ("tulay", "nhom1", "1"):
            group = "Tự lấy"
        location = clean(p.get("location", old.get("Mã vị trí", ""))).upper().replace(" ", "")
        if not name:
            return None, "Nhập tên vật tư."
        if group not in INV_GROUPS:
            return None, "Nhóm chỉ nhận Báo KTV hoặc Tự lấy."
        if location and not location_root(location):
            return None, f"Mã vị trí {location} không đúng dạng (vd T1-N2-04, K2-N1, D4)."
        values = {}
        for field, header in (("std", "SL chuẩn"), ("min", "SL tối thiểu")):
            raw = clean(p.get(field, old.get(header, "")))
            if raw and to_int(raw) is None:
                return None, f"{header} phải là số."
            values[header] = "" if not raw else to_int(raw)
        row = dict(old)
        row.update({
            "Tên vật tư": name, "Nhóm": group, "Mã vị trí": location,
            "Đơn vị": clean(p.get("unit", old.get("Đơn vị", ""))),
            "SL chuẩn": values["SL chuẩn"], "SL tối thiểu": values["SL tối thiểu"],
            "Ghi chú": clean(p.get("note", old.get("Ghi chú", ""))),
            "Trạng thái": clean(p.get("status", old.get("Trạng thái", ""))) or INV_ACTIVE,
            "Cập nhật lúc": now_text(), "Người cập nhật": staff_name(),
        })
        if row["Trạng thái"] not in (INV_ACTIVE, INV_RETIRED):
            row["Trạng thái"] = INV_ACTIVE
        return row, ""

    @app.route("/api/inventory")
    def api_inventory():
        denied = require_admin()
        if denied:
            return denied
        try:
            return jsonify({"success": True, "data": inventory_view(), "groups": INV_GROUPS,
                            "locations": STORAGE_LOCATIONS, "count_due_days": COUNT_DUE_DAYS})
        except Exception as e:
            print("API INVENTORY ERROR:", e)
            return jsonify({"success": False, "data": [], "error": str(e)}), 500

    @app.route("/api/inventory", methods=["POST"])
    def api_inventory_save():
        denied = require_admin()
        if denied:
            return denied
        p = request.get_json(silent=True) or {}
        item_id = clean(p.get("id")).upper()
        try:
            with _lock:
                sheet, items = read(INV_SHEET, INV_HEADERS, fresh=True)
                if item_id:
                    old = next((x for x in items if x["Mã vật tư"].upper() == item_id), None)
                    if not old:
                        return jsonify({"success": False, "error": "Không tìm thấy vật tư."}), 404
                    row, error = item_row(p, old)
                    if error:
                        return jsonify({"success": False, "error": error}), 400
                    write_row(sheet, old["_row"], INV_HEADERS, row)
                    message = f"Đã cập nhật {item_id}."
                else:
                    row, error = item_row(p)
                    if error:
                        return jsonify({"success": False, "error": error}), 400
                    key = (norm_name(row["Tên vật tư"]), row["Mã vị trí"])
                    if any((norm_name(x["Tên vật tư"]), x["Mã vị trí"].upper()) == key for x in items):
                        return jsonify({"success": False, "error": "Món này đã có trong danh mục ở cùng vị trí."}), 409
                    item_id = row["Mã vật tư"] = next_item_code(items)
                    sheet.append_row([row.get(k, "") for k in INV_HEADERS], value_input_option="RAW")
                    message = f"Đã thêm {row['Tên vật tư']} ({item_id})."
                invalidate(INV_SHEET)
            return jsonify({"success": True, "id": item_id, "message": message})
        except Exception as e:
            print("API INVENTORY SAVE ERROR:", e)
            return jsonify({"success": False, "error": str(e)}), 500

    @app.route("/api/inventory/import", methods=["POST"])
    def api_inventory_import():
        """Thêm nhiều món một lần (dán từ Excel hoặc danh sách món Báo KTV của SOP-03)."""
        denied = require_admin()
        if denied:
            return denied
        rows = (request.get_json(silent=True) or {}).get("rows") or []
        if not isinstance(rows, list) or not rows:
            return jsonify({"success": False, "error": "Không có dòng nào để thêm."}), 400
        if len(rows) > 500:
            return jsonify({"success": False, "error": "Mỗi lần thêm tối đa 500 dòng."}), 400
        try:
            with _lock:
                sheet, items = read(INV_SHEET, INV_HEADERS, fresh=True)
                existing = {(norm_name(x["Tên vật tư"]), x["Mã vị trí"].upper()) for x in items}
                new_rows, skipped, errors = [], 0, []
                for index, p in enumerate(rows, start=1):
                    if not isinstance(p, dict):
                        continue
                    row, error = item_row(p)
                    if error:
                        errors.append(f"dòng {index}: {error}")
                        continue
                    key = (norm_name(row["Tên vật tư"]), row["Mã vị trí"])
                    if key in existing:
                        skipped += 1
                        continue
                    existing.add(key)
                    row["Mã vật tư"] = next_item_code(items + [{"Mã vật tư": r[0]} for r in new_rows])
                    new_rows.append([row.get(k, "") for k in INV_HEADERS])
                if errors:
                    return jsonify({"success": False, "error": "Chưa thêm gì. Sửa lại: " + "; ".join(errors[:8])
                                    + (f" (và {len(errors) - 8} dòng khác)" if len(errors) > 8 else "")}), 400
                if new_rows:
                    sheet.append_rows(new_rows, value_input_option="RAW")
                    invalidate(INV_SHEET)
            message = f"Đã thêm {len(new_rows)} món."
            if skipped:
                message += f" Bỏ qua {skipped} món đã có ở cùng vị trí."
            return jsonify({"success": True, "added": len(new_rows), "skipped": skipped, "message": message})
        except Exception as e:
            print("API INVENTORY IMPORT ERROR:", e)
            return jsonify({"success": False, "error": str(e)}), 500

    @app.route("/api/inventory/<item_id>/count", methods=["POST"])
    def api_inventory_count(item_id):
        denied = require_admin()
        if denied:
            return denied
        p = request.get_json(silent=True) or {}
        qty, broken = to_int(p.get("qty")), to_int(p.get("broken")) or 0
        if qty is None:
            return jsonify({"success": False, "error": "Nhập số lượng đếm được."}), 400
        if broken > qty:
            return jsonify({"success": False, "error": "Số hỏng không lớn hơn số đếm được."}), 400
        try:
            item = next((x for x in inventory_view() if x["id"] == clean(item_id).upper()), None)
            if not item:
                return jsonify({"success": False, "error": "Không tìm thấy vật tư."}), 404
            with _lock:
                sheet = get_sheet(COUNT_SHEET, COUNT_HEADERS)
                count_id = "KK-" + datetime.now().strftime("%y%m%d") + "-" + uuid.uuid4().hex[:4].upper()
                row = {"Mã kiểm kê": count_id, "Thời gian": now_text(), "Mã vật tư": item["id"],
                       "Tên vật tư": item["name"], "Mã vị trí": item["location"], "SL thực tế": qty,
                       "Số hỏng": broken, "Ghi chú": clean(p.get("note")), "KTV": staff_name()}
                sheet.append_row([row.get(k, "") for k in COUNT_HEADERS], value_input_option="RAW")
                invalidate(COUNT_SHEET)
            message = f"Đã ghi kiểm kê {item['name']}: {qty}" + (f" (hỏng {broken})" if broken else "") + "."
            return jsonify({"success": True, "id": count_id, "message": message})
        except Exception as e:
            print("API INVENTORY COUNT ERROR:", e)
            return jsonify({"success": False, "error": str(e)}), 500

    @app.route("/api/inventory/<item_id>/history")
    def api_inventory_history(item_id):
        denied = require_admin()
        if denied:
            return denied
        try:
            code = clean(item_id).upper()
            item = next((x for x in inventory_view() if x["id"] == code), None)
            if not item:
                return jsonify({"success": False, "error": "Không tìm thấy vật tư."}), 404
            _, counts = read(COUNT_SHEET, COUNT_HEADERS)
            _, loans = read(LOAN_SHEET, LOAN_HEADERS)
            name = norm_name(item["name"])
            count_rows = [{"at": c["Thời gian"], "qty": c["SL thực tế"], "broken": c["Số hỏng"], "note": c["Ghi chú"],
                           "by": c["KTV"]} for c in counts if c["Mã vật tư"].upper() == code]
            loan_rows = [loan_public(l) for l in loans
                         if l["Mã vật tư"].upper() == code or (not l["Mã vật tư"] and norm_name(l["Món / thiết bị"]) == name)]
            count_rows.sort(key=lambda x: (parse_dt(x["at"]) or datetime.min), reverse=True)
            loan_rows.sort(key=lambda x: (parse_dt(x["borrowed_at"]) or datetime.min), reverse=True)
            return jsonify({"success": True, "item": item, "counts": count_rows[:30], "loans": loan_rows[:30]})
        except Exception as e:
            print("API INVENTORY HISTORY ERROR:", e)
            return jsonify({"success": False, "error": str(e)}), 500

    # =====================================================
    # CHÈN lab-scope.js VÀO MỌI TRANG (chế độ chỉ hiện Lab 1)
    # =====================================================
    @app.after_request
    def inject_lab_scope(response):
        try:
            if response.mimetype == "text/html" and not response.direct_passthrough:
                body = response.get_data(as_text=True)
                tag = '<script src="/static/lab-scope.js"></script>'
                if "</head>" in body and tag not in body:
                    response.set_data(body.replace("</head>", tag + "\n</head>", 1))
        except Exception as e:
            print("LAB SCOPE INJECT ERROR:", e)
        return response

    return app
