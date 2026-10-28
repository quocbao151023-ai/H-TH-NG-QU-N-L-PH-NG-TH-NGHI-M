from flask import Flask, render_template, jsonify, request, redirect, url_for, session, send_file
from datetime import datetime, timedelta
import csv
import io
import re
import threading
import time
import uuid
import unicodedata
import gspread
from google.oauth2.service_account import Credentials

app = Flask(__name__)
app.secret_key = "lab-system-secret-key"

# Giới hạn file tải lên (danh sách lớp) tối đa 5 MB
app.config["MAX_CONTENT_LENGTH"] = 5 * 1024 * 1024

# =========================
# TÀI KHOẢN ĐĂNG NHẬP
# =========================

TECH_EMAIL = "Quocbao151023@gmail.com"
TECH_PASSWORD = "p4ssw0rd"
TECH_NAME = "Quốc Bảo"


# =========================
# TÀI KHOẢN GIẢNG VIÊN
# Thêm / sửa giảng viên tại đây:
#   "email đăng nhập": {"password": "...", "name": "Họ tên hiển thị"}
# =========================

LECTURER_ACCOUNTS = {
    "giangvien@vlu.edu.vn": {
        "password": "giangvien123",
        "name": "Giảng viên Demo"
    },
}



# =========================
# GOOGLE SHEET CONFIG
# =========================

GOOGLE_CREDENTIALS_FILE = "key.json"

BOOKING_SPREADSHEET_ID = "16eguXNHfCzbTq6olAQaMQBAylkqNR3gjVMmtnAkSBpE"
HANDOVER_SPREADSHEET_ID = "1nmIwGAMIFnuAfv-EWGG9yubxjifodxvsX2Iyzx0fJRU"

BOOKING_SHEET_NAME = "LAB_BOOKING"
HANDOVER_SHEET_NAME = "LAB_HANDOVER"
ATTENDANCE_SHEET_NAME = "LAB_ATTENDANCE"

# Sheet lưu danh sách lớp do giảng viên tải lên (tự tạo trong file Booking)
ROSTER_SHEET_NAME = "LECTURER_ROSTER"
ROSTER_HEADERS = [
    "Dấu thời gian",
    "Email giảng viên",
    "Tên giảng viên",
    "Mã lớp",
    "Môn học",
    "MSSV",
    "Họ tên",
    "Nhóm",
    "Đề tài"
]


# =========================
# GOOGLE SHEET CONNECT
# =========================

def get_google_client(readonly=False):
    if readonly:
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets.readonly",
            "https://www.googleapis.com/auth/drive.readonly"
        ]
    else:
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]

    credentials = Credentials.from_service_account_file(
        GOOGLE_CREDENTIALS_FILE,
        scopes=scopes
    )

    return gspread.authorize(credentials)


client = get_google_client(readonly=False)


def _get_sheet_rows_uncached(spreadsheet_id, sheet_name):
    google_client = get_google_client(readonly=True)
    spreadsheet = google_client.open_by_key(spreadsheet_id)
    worksheet = spreadsheet.worksheet(sheet_name)
    return worksheet.get_all_records()


# =========================
# BỘ NHỚ ĐỆM ĐỌC GOOGLE SHEET
# Google giới hạn 60 lượt đọc/phút. Mỗi lần đọc 1 sheet tốn ~3 lượt
# (mở file + mở trang tính + đọc dữ liệu), nên giữ kết quả trong 30 giây.
# Ghi dữ liệu xong thì xóa bộ nhớ đệm tương ứng ngay để luôn thấy dữ liệu mới.
# Nếu Google báo quá giới hạn (lỗi 429) mà đã có dữ liệu cũ -> dùng tạm dữ liệu cũ.
# =========================

SHEET_CACHE_SECONDS = 30

_sheet_cache = {}
_sheet_cache_lock = threading.Lock()


def cached_call(key, loader, ttl=SHEET_CACHE_SECONDS):
    now = time.time()

    with _sheet_cache_lock:
        hit = _sheet_cache.get(key)

    if hit and now - hit[0] < ttl:
        return hit[1]

    try:
        value = loader()
    except Exception as e:
        if hit:
            print("SHEET CACHE: dùng tạm dữ liệu cũ vì lỗi:", e)
            return hit[1]

        if "429" in str(e) or "Quota exceeded" in str(e):
            raise RuntimeError(
                "Google Sheets đang giới hạn số lượt đọc (tối đa 60 lượt/phút). "
                "Vui lòng thử lại sau khoảng 1 phút."
            ) from e

        raise

    with _sheet_cache_lock:
        _sheet_cache[key] = (now, value)

    return value


def invalidate_cache(*names):
    """Xóa bộ nhớ đệm có phần tử đầu của khóa nằm trong names."""
    with _sheet_cache_lock:
        for key in list(_sheet_cache):
            if key[0] in names or (len(key) > 1 and key[1] in names):
                del _sheet_cache[key]


class CachedSpreadsheet:
    """Giữ file Google Sheet đã mở và các trang tính đã tìm (không mở lại mỗi lần)."""

    def __init__(self, spreadsheet):
        self._spreadsheet = spreadsheet
        self._worksheets = {}
        self._lock = threading.Lock()

    def worksheet(self, title):
        with self._lock:
            sheet = self._worksheets.get(title)

        if sheet is None:
            sheet = self._spreadsheet.worksheet(title)
            with self._lock:
                self._worksheets[title] = sheet

        return sheet

    def add_worksheet(self, title, rows, cols, **kwargs):
        sheet = self._spreadsheet.add_worksheet(title=title, rows=rows, cols=cols, **kwargs)
        with self._lock:
            self._worksheets[title] = sheet
        return sheet

    def __getattr__(self, name):
        return getattr(self._spreadsheet, name)


_cached_spreadsheets = {}


def sheet_values_cached(sheet, name):
    """get_all_values() có bộ nhớ đệm (dùng cho các sheet hệ thống tự tạo)."""
    return cached_call(("values", name), sheet.get_all_values)


def get_sheet_rows(spreadsheet_id, sheet_name):
    rows = cached_call(
        ("rows", spreadsheet_id, sheet_name),
        lambda: _get_sheet_rows_uncached(spreadsheet_id, sheet_name)
    )
    return [dict(row) for row in rows]


def get_lab_booking_rows():
    rows = get_sheet_rows(
        BOOKING_SPREADSHEET_ID,
        BOOKING_SHEET_NAME
    )

    # + lượt đăng ký trên cổng sinh viên (bỏ qua mã đã có trong dữ liệu form)
    existing = {clean_value(row.get("ID_KEY")) for row in rows if clean_value(row.get("ID_KEY"))}
    return rows + portal_booking_rows(existing)


def get_lab_handover_rows():
    rows = get_sheet_rows(
        HANDOVER_SPREADSHEET_ID,
        HANDOVER_SHEET_NAME
    )

    # + phiếu bàn giao trên cổng sinh viên
    return rows + portal_handover_rows()


def get_booking_spreadsheet_for_write():
    spreadsheet = _cached_spreadsheets.get(BOOKING_SPREADSHEET_ID)

    if spreadsheet is None:
        spreadsheet = CachedSpreadsheet(client.open_by_key(BOOKING_SPREADSHEET_ID))
        _cached_spreadsheets[BOOKING_SPREADSHEET_ID] = spreadsheet

    return spreadsheet

def get_or_create_attendance_sheet():
    spreadsheet = get_booking_spreadsheet_for_write()

    try:
        sheet = spreadsheet.worksheet(ATTENDANCE_SHEET_NAME)
    except Exception:
        sheet = spreadsheet.add_worksheet(
            title=ATTENDANCE_SHEET_NAME,
            rows=1000,
            cols=10
        )

        sheet.append_row([
            "Dấu thời gian",
            "Mã sinh viên",
            "Họ tên",
            "Mã đăng ký",
            "Lab",
            "Ngày",
            "Ca",
            "Loại",
            "Trạng thái điểm danh",
            "Người điểm danh"
        ])

    return sheet


def get_attendance_rows():
    try:
        rows = cached_call(
            ("attendance", ATTENDANCE_SHEET_NAME),
            lambda: get_or_create_attendance_sheet().get_all_records()
        )
        return [dict(row) for row in rows]
    except Exception:
        return []


def make_attendance_key(booking_id, student_id, name, lab, date, shift):
    return "|".join([
        clean_value(booking_id),
        clean_value(student_id),
        clean_value(name).lower(),
        clean_value(lab),
        clean_value(date),
        clean_value(shift)
    ])

# =========================
# HELPER FUNCTIONS
# =========================

def clean_value(value):
    if value is None:
        return ""
    return str(value).strip()


def remove_vietnamese_accents(text):
    text = str(text or "")
    text = unicodedata.normalize("NFD", text)
    text = "".join(
        char for char in text
        if unicodedata.category(char) != "Mn"
    )
    text = text.replace("đ", "d").replace("Đ", "D")
    return text


def normalize_text(value):
    return remove_vietnamese_accents(clean_value(value)).lower().replace(" ", "")


def normalize_yes_no_py(value):
    text = remove_vietnamese_accents(clean_value(value)).lower()

    if text in ["co", "yes", "true", "1"]:
        return "CO"

    if text in ["khong", "no", "false", "0"]:
        return "KHONG"

    return text.upper()


def get_value_by_possible_keys(row, keys):
    for key in keys:
        if key in row:
            value = clean_value(row.get(key))
            if value:
                return value

    normalized_row = {
        clean_value(k): v
        for k, v in row.items()
    }

    for key in keys:
        key_clean = clean_value(key)
        if key_clean in normalized_row:
            value = clean_value(normalized_row.get(key_clean))
            if value:
                return value

    row_by_normalized_key = {
        normalize_text(k): v
        for k, v in row.items()
    }

    for key in keys:
        key_norm = normalize_text(key)
        if key_norm in row_by_normalized_key:
            value = clean_value(row_by_normalized_key.get(key_norm))
            if value:
                return value

    return ""


def extract_lab_from_id_key(id_key):
    if not id_key:
        return ""

    match = re.match(r"^L([1-3])-?C[1-4]", str(id_key).strip().upper())

    return f"Lab {match.group(1)}" if match else ""


def extract_shift_from_id_key(id_key):
    if not id_key:
        return ""

    match = re.match(r"^L[1-3]-?C([1-4])", str(id_key).strip().upper())

    return f"Ca {match.group(1)}" if match else ""


def extract_lab_from_slot_key(slot_key):
    if not slot_key:
        return ""

    parts = str(slot_key).split("|")

    for part in parts:
        part = part.strip()
        if part.lower().startswith("lab"):
            return part

    return ""


def extract_shift_from_slot_key(slot_key):
    if not slot_key:
        return ""

    parts = str(slot_key).split("|")

    for part in parts:
        part = part.strip()
        if part.lower().startswith("ca"):
            return part

    return ""


def format_date_for_calendar(raw_date):
    raw_date = clean_value(raw_date)

    if not raw_date:
        return ""

    formats = [
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%m/%d/%Y",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(raw_date, fmt).strftime("%Y-%m-%d")
        except Exception:
            pass

    return ""


def parse_notification_date(raw_date):
    raw_date = clean_value(raw_date)

    if not raw_date:
        return None

    formats = [
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%m/%d/%Y",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(raw_date, fmt)
        except Exception:
            pass

    return None


def is_within_notification_window(raw_date):
    date_obj = parse_notification_date(raw_date)

    if not date_obj:
        return False

    now = datetime.now()
    start_date = now - timedelta(days=7)
    end_date = now + timedelta(days=7)

    return start_date <= date_obj <= end_date


def parse_group_members(member_text):
    members = []

    member_text = clean_value(member_text)

    if not member_text:
        return members

    lines = member_text.replace(";", "\n").replace(",", "\n").split("\n")

    for line in lines:
        text = line.strip()

        if not text:
            continue

        if "." in text:
            first_part = text.split(".", 1)[0].strip()
            if first_part.isdigit():
                text = text.split(".", 1)[1].strip()

        if "–" in text:
            parts = text.split("–")
        elif "-" in text:
            parts = text.split("-")
        else:
            parts = [text]

        name = parts[0].strip()
        student_id = ""

        if len(parts) > 1:
            student_id = parts[1].strip()

        if name or student_id:
            members.append({
                "name": name,
                "student_id": student_id
            })

    return members


# =========================
# PAGE ROUTES + LOGIN
# =========================

@app.route("/", methods=["GET", "POST"])
def login():
    error = None

    # Khi mở link gốc 127.0.0.1:5000/
    # luôn hiện giao diện đăng nhập trước
    if request.method == "GET":
        session.clear()
        return render_template("login.html", error=error)

    # Khi bấm đăng nhập
    email = request.form.get("email")
    password = request.form.get("password")

    if email == TECH_EMAIL and password == TECH_PASSWORD:
        session["user"] = TECH_NAME
        session["role"] = "admin"
        session["email"] = TECH_EMAIL
        return redirect(url_for("home"))

    lecturer = LECTURER_ACCOUNTS.get(clean_value(email).lower())

    if lecturer and password == lecturer.get("password"):
        session["user"] = lecturer.get("name") or email
        session["role"] = "lecturer"
        session["email"] = clean_value(email).lower()
        return redirect(url_for("teacher"))

    error = "Sai email hoặc mật khẩu"
    return render_template("login.html", error=error)


@app.route("/dashboard")
def dashboard():
    if "user" not in session:
        return redirect(url_for("login"))

    if session.get("role") == "lecturer":
        return redirect(url_for("teacher"))

    return render_template(
        "dashboard.html",
        user=session["user"]
    )


@app.route("/attendance")
def attendance():
    if "user" not in session:
        return redirect(url_for("login"))

    if session.get("role") == "lecturer":
        return redirect(url_for("teacher"))

    return render_template(
        "attendance.html",
        user=session["user"]
    )


@app.route("/cleaning-report")
def cleaning_report():
    if "user" not in session:
        return redirect(url_for("login"))

    if session.get("role") == "lecturer":
        return redirect(url_for("teacher"))

    return render_template(
        "cleaning_report.html",
        user=session["user"]
    )


@app.route("/lab-performance")
def lab_performance():
    if "user" not in session:
        return redirect(url_for("login"))

    if session.get("role") == "lecturer":
        return redirect(url_for("teacher"))

    return render_template(
        "lab_performance.html",
        user=session["user"]
    )


@app.route("/teacher")
def teacher():
    if "user" not in session:
        return redirect(url_for("login"))

    return render_template(
        "teacher.html",
        user=session["user"]
    )


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# =========================
# API THÔNG BÁO TỪ SHEET
# =========================

def build_notifications():
    """Danh sách thông báo (đăng ký, yêu cầu KTV, vấn đề bàn giao) trong khoảng ±7 ngày.
    Dùng chung cho chuông thông báo và Trang chủ."""
    bookings = get_lab_booking_rows()
    handovers = get_lab_handover_rows()

    notifications = []

    for row in bookings:
        status = get_value_by_possible_keys(row, [
            "TRẠNG THÁI",
            "TRANG THAI"
        ])

        if status != "ĐÃ_XÁC_NHẬN":
            continue

        id_key = get_value_by_possible_keys(row, [
            "ID_KEY"
        ])

        name = get_value_by_possible_keys(row, [
            "1. Họ và tên của bạn là:",
            "1. Họ và tên của bạn là: ",
            "Họ và tên của bạn là:",
            "Họ và tên"
        ])

        slot_key = get_value_by_possible_keys(row, [
            "SLOT_KEY"
        ])

        lab = (
            get_value_by_possible_keys(row, [
                "6. Phòng lab được lựa chọn là:",
                "6. Phòng Lab được lựa chọn là:",
                "Phòng lab được lựa chọn là:",
                "Phòng Lab",
                "Lab"
            ])
            or extract_lab_from_slot_key(slot_key)
            or extract_lab_from_id_key(id_key)
        )

        shift = (
            get_value_by_possible_keys(row, [
                "9. Ca bạn đăng ký sử dụng là:",
                "9. Ca bạn đăng ký sử dụng là: ",
                "Ca bạn đăng ký sử dụng là:",
                "Ca"
            ])
            or extract_shift_from_slot_key(slot_key)
            or extract_shift_from_id_key(id_key)
        )

        use_date = get_value_by_possible_keys(row, [
            "8. Ngày bạn đăng ký sử dụng phòng là:",
            "Ngày bạn đăng ký sử dụng phòng là:",
            "Ngày đăng ký",
            "Ngày sử dụng"
        ])

        support = get_value_by_possible_keys(row, [
            "10. Bạn có cần hỗ trợ kỹ thuật không?",
            "11. Bạn có cần hỗ trợ kỹ thuật không?",
            "Cần hỗ trợ kỹ thuật",
            "Hỗ trợ kỹ thuật"
        ])

        if not id_key:
            continue

        if not is_within_notification_window(use_date):
            continue

        support_text = remove_vietnamese_accents(str(support)).lower()
        need_support = (
            "co" in support_text
            or "có" in str(support).lower()
        )

        if need_support:
            notifications.append({
                "notification_id": f"booking-support-{id_key}",
                "type": "booking_support",
                "title": "THÔNG BÁO ĐĂNG KÝ - YÊU CẦU KTV",
                "message": f"{name} đã đăng ký {lab} - {shift}",
                "id_key": id_key,
                "date": use_date,
                "level": "warning"
            })
        else:
            notifications.append({
                "notification_id": f"booking-{id_key}",
                "type": "booking",
                "title": "THÔNG BÁO ĐĂNG KÝ PHÒNG LAB",
                "message": f"{name} đã đăng ký {lab} - {shift}",
                "id_key": id_key,
                "date": use_date,
                "level": "info"
            })

    for row in handovers:
        id_key = get_value_by_possible_keys(row, [
            "1. ID_KEY của bạn là:",
            "ID_KEY",
            "ID_KEY của bạn là:"
        ])

        student_name = get_value_by_possible_keys(row, [
            "3. Họ và tên của bạn là:",
            "5. Họ và tên của bạn là:",
            "Họ và tên"
        ])

        process_status = get_value_by_possible_keys(row, [
            "PROCESS_STATUS"
        ])

        error_message = get_value_by_possible_keys(row, [
            "ERROR_MESSAGE"
        ])

        abnormal = get_value_by_possible_keys(row, [
            "10. Phòng Lab có phát sinh vấn đề bất thường trong quá trình sử dụng không?",
            "Phòng Lab có phát sinh vấn đề bất thường trong quá trình sử dụng không?"
        ])

        handover_date = get_value_by_possible_keys(row, [
            "Dấu thời gian",
            "UPDATED_AT"
        ])

        abnormal_text = remove_vietnamese_accents(str(abnormal)).lower()

        has_problem = False

        if process_status == "ERROR":
            has_problem = True

        if error_message:
            has_problem = True

        if "co" in abnormal_text or "có" in str(abnormal).lower():
            has_problem = True

        if has_problem and id_key and is_within_notification_window(handover_date):
            notifications.append({
                "notification_id": f"handover-{id_key}",
                "type": "handover_problem",
                "title": "CÓ VẤN ĐỀ BÀN GIAO",
                "message": error_message or f"{student_name} báo có vấn đề bàn giao",
                "id_key": id_key,
                "date": handover_date,
                "level": "danger"
            })

    notifications.sort(
        key=lambda item: parse_notification_date(item.get("date")) or datetime.min,
        reverse=True
    )

    return merge_multi_shift(notifications, "notification_id")


@app.route("/api/notifications")
def api_notifications():
    try:
        notifications = build_notifications()

        notifications = notifications[:30]

        return jsonify({
            "success": True,
            "count": len(notifications),
            "data": notifications
        })

    except Exception as e:
        print("API NOTIFICATIONS ERROR:", e)

        return jsonify({
            "success": False,
            "count": 0,
            "data": [],
            "error": str(e)
        }), 500


# =========================
# API LỊCH PHÒNG LAB
# =========================

@app.route("/api/bookings")
def get_bookings():
    try:
        rows = get_lab_booking_rows()
        bookings = []

        for row in rows:
            status = get_value_by_possible_keys(row, [
                "TRẠNG THÁI",
                "TRANG THAI"
            ])

            if status != "ĐÃ_XÁC_NHẬN":
                continue

            raw_date = get_value_by_possible_keys(row, [
                "8. Ngày bạn đăng ký sử dụng phòng là:",
                "Ngày bạn đăng ký sử dụng phòng là:",
                "Ngày đăng ký"
            ])

            formatted_date = format_date_for_calendar(raw_date)

            if not formatted_date:
                continue

            id_key = get_value_by_possible_keys(row, [
                "ID_KEY",
                "1. ID_KEY của bạn là:"
            ])

            slot_key = get_value_by_possible_keys(row, [
                "SLOT_KEY"
            ])

            lab = (
                get_value_by_possible_keys(row, [
                    "6. Phòng lab được lựa chọn là:",
                    "6. Phòng Lab được lựa chọn là:",
                    "Phòng lab được lựa chọn là:",
                    "Phòng Lab",
                    "Lab"
                ])
                or extract_lab_from_slot_key(slot_key)
                or extract_lab_from_id_key(id_key)
            )

            shift = (
                get_value_by_possible_keys(row, [
                    "9. Ca bạn đăng ký sử dụng là: ",
                    "9. Ca bạn đăng ký sử dụng là:",
                    "Ca bạn đăng ký sử dụng là:",
                    "Ca"
                ])
                or extract_shift_from_slot_key(slot_key)
                or extract_shift_from_id_key(id_key)
            )

            name = get_value_by_possible_keys(row, [
                "1. Họ và tên của bạn là: ",
                "1. Họ và tên của bạn là:",
                "Họ và tên của bạn là:",
                "Họ và tên"
            ])

            bookings.append({
                "name": name,
                "lab": lab,
                "shift": shift,
                "date": formatted_date,
                "status": status,
                "idKey": id_key,
                **booking_detail_fields(row)
            })

        return jsonify(bookings)

    except Exception as e:
        print("API BOOKINGS ERROR:", e)

        return jsonify({
            "success": False,
            "error": str(e)
        }), 500


# =========================
# API THEO DÕI ĐIỂM DANH
# =========================

def build_attendance_records():
    """Danh sách sinh viên đã đăng ký (kể cả thành viên nhóm) + trạng thái điểm danh.
    Dùng chung cho trang Theo dõi ra vào và trang Theo dõi theo lớp."""
    rows = get_lab_booking_rows()
    attendance_rows = get_attendance_rows()

    latest_status = {}

    for att in attendance_rows:
        key = make_attendance_key(
            get_value_by_possible_keys(att, ["Mã đăng ký"]),
            get_value_by_possible_keys(att, ["Mã sinh viên"]),
            get_value_by_possible_keys(att, ["Họ tên"]),
            get_value_by_possible_keys(att, ["Lab"]),
            get_value_by_possible_keys(att, ["Ngày"]),
            get_value_by_possible_keys(att, ["Ca"])
        )

        status = get_value_by_possible_keys(att, ["Trạng thái điểm danh"])

        if key and status:
            latest_status[key] = status

    data = []

    for row in rows:
        status = get_value_by_possible_keys(row, [
            "TRẠNG THÁI",
            "TRANG THAI"
        ])

        if status != "ĐÃ_XÁC_NHẬN":
            continue

        id_key = get_value_by_possible_keys(row, ["ID_KEY"])
        slot_key = get_value_by_possible_keys(row, ["SLOT_KEY"])

        register_type = get_value_by_possible_keys(row, [
            "5. Bạn đăng ký với tư cách là:",
            "Bạn đăng ký với tư cách là:"
        ])

        name = get_value_by_possible_keys(row, [
            "1. Họ và tên của bạn là: ",
            "1. Họ và tên của bạn là:",
            "Họ và tên của bạn là:",
            "Họ và tên"
        ])

        student_id = get_value_by_possible_keys(row, [
            "2. Mã số sinh viên của bạn là:",
            "Mã số sinh viên của bạn là:",
            "MSSV"
        ])

        lab = (
            get_value_by_possible_keys(row, [
                "6. Phòng lab được lựa chọn là:",
                "6. Phòng Lab được lựa chọn là:",
                "Phòng lab được lựa chọn là:",
                "Phòng",
                "Lab"
            ])
            or extract_lab_from_slot_key(slot_key)
            or extract_lab_from_id_key(id_key)
        )

        date = get_value_by_possible_keys(row, [
            "8. Ngày bạn đăng ký sử dụng phòng là:",
            "Ngày bạn đăng ký sử dụng phòng là:",
            "Ngày đăng ký"
        ])

        shift = (
            get_value_by_possible_keys(row, [
                "9. Ca bạn đăng ký sử dụng là: ",
                "9. Ca bạn đăng ký sử dụng là:",
                "Ca bạn đăng ký sử dụng là:",
                "Ca"
            ])
            or extract_shift_from_slot_key(slot_key)
            or extract_shift_from_id_key(id_key)
        )

        member_list = get_value_by_possible_keys(row, [
            "5.2. Danh sách thành viên nhóm là: (Họ tên 1 - MSSV 1; Họ tên 2 - MSSV 2; ...) ",
            "5.2. Danh sách thành viên nhóm là: (Họ tên 1 - MSSV 1; Họ tên 2 - MSSV 2; ...)",
            "5.2. Danh sách thành viên nhóm là:",
            "Danh sách thành viên nhóm là:"
        ])

        if not name and not student_id:
            continue

        main_key = make_attendance_key(id_key, student_id, name, lab, date, shift)

        data.append({
            "booking_id": id_key,
            "register_type": register_type,
            "name": name,
            "student_id": student_id,
            "lab": lab,
            "date": date,
            "shift": shift,
            "role": "Người đăng ký chính",
            "raw_members": member_list,
            "attendance_status": latest_status.get(main_key, "Chưa điểm danh")
        })

        if register_type == "Nhóm" and member_list:
            members = parse_group_members(member_list)

            for member in members:
                member_key = make_attendance_key(
                    id_key,
                    member["student_id"],
                    member["name"],
                    lab,
                    date,
                    shift
                )

                data.append({
                    "booking_id": id_key,
                    "register_type": "Nhóm",
                    "name": member["name"],
                    "student_id": member["student_id"],
                    "lab": lab,
                    "date": date,
                    "shift": shift,
                    "role": "Thành viên nhóm",
                    "raw_members": member_list,
                    "attendance_status": latest_status.get(member_key, "Chưa điểm danh")
                })

    return data


@app.route("/api/attendance-bookings")
def attendance_bookings():
    try:
        data = attach_handover_flags(build_attendance_records())

        return jsonify({
            "success": True,
            "data": data
        })

    except Exception as e:
        print("API ATTENDANCE BOOKINGS ERROR:", e)

        return jsonify({
            "success": False,
            "error": str(e)
        }), 500
# =========================
# API LƯU TRẠNG THÁI ĐIỂM DANH
# =========================

@app.route("/api/save-attendance-status", methods=["POST"])
def save_attendance_status():
    try:
        data = request.get_json()
        spreadsheet = get_booking_spreadsheet_for_write()

        try:
            sheet = spreadsheet.worksheet(ATTENDANCE_SHEET_NAME)
        except Exception:
            sheet = spreadsheet.add_worksheet(
                title=ATTENDANCE_SHEET_NAME,
                rows=1000,
                cols=10
            )

            sheet.append_row([
                "Dấu thời gian",
                "Mã sinh viên",
                "Họ tên",
                "Mã đăng ký",
                "Lab",
                "Ngày",
                "Ca",
                "Loại",
                "Trạng thái điểm danh",
                "Người điểm danh"
            ])

        now = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        sheet.append_row([
            now,
            data.get("student_id", ""),
            data.get("name", ""),
            data.get("booking_id", ""),
            data.get("lab", ""),
            data.get("date", ""),
            data.get("shift", ""),
            data.get("role", ""),
            data.get("attendance_status", ""),
            staff_on_duty_name(data.get("date", ""), data.get("shift", ""), data.get("lab", ""))
        ])

        invalidate_cache("attendance")

        return jsonify({
            "success": True
        })

    except Exception as e:
        print("API SAVE ATTENDANCE ERROR:", e)

        return jsonify({
            "success": False,
            "error": str(e)
        }), 500


# =========================
# API BÁO CÁO VỆ SINH / BÀN GIAO
# =========================

def build_cleaning_records():
    """Trạng thái bàn giao của từng lượt đăng ký.
    Dùng chung cho trang Vệ sinh & bàn giao và file Excel của giảng viên."""
    bookings = get_lab_booking_rows()
    handovers = get_lab_handover_rows()

    handover_map = {}

    for handover in handovers:
        id_key = get_value_by_possible_keys(handover, [
            "1. ID_KEY của bạn là:",
            "ID_KEY",
            "ID_KEY của bạn là:"
        ])

        if id_key:
            handover_map[id_key] = handover

    report_data = []

    for booking in bookings:
        id_key = get_value_by_possible_keys(booking, [
            "ID_KEY",
            "1. ID_KEY của bạn là:"
        ])

        if not id_key:
            continue

        status_booking = get_value_by_possible_keys(booking, [
            "TRẠNG THÁI"
        ])

        if status_booking and status_booking != "ĐÃ_XÁC_NHẬN":
            continue

        handover = handover_map.get(id_key)

        slot_key_booking = get_value_by_possible_keys(booking, [
            "SLOT_KEY"
        ])

        lab = (
            get_value_by_possible_keys(booking, [
                "6. Phòng lab được lựa chọn là:",
                "6. Phòng Lab được lựa chọn là:",
                "9. Phòng Lab bạn đã sử dụng là:",
                "Phòng",
                "Lab"
            ])
            or extract_lab_from_slot_key(slot_key_booking)
            or extract_lab_from_id_key(id_key)
        )

        shift = (
            get_value_by_possible_keys(booking, [
                "9. Ca bạn đăng ký sử dụng là: ",
                "9. Ca bạn đăng ký sử dụng là:",
                "8. Ca bạn đăng ký sử dụng là:",
                "Ca"
            ])
            or extract_shift_from_slot_key(slot_key_booking)
            or extract_shift_from_id_key(id_key)
        )

        booking_date = get_value_by_possible_keys(booking, [
            "8. Ngày bạn đăng ký sử dụng phòng là:",
            "Ngày đăng ký",
            "Dấu thời gian"
        ])

        if handover:
            process_status = get_value_by_possible_keys(handover, [
                "PROCESS_STATUS"
            ])

            if process_status == "DONE":
                status = "Đã bàn giao"
            else:
                status = "Cần kiểm tra"

            slot_key_handover = get_value_by_possible_keys(handover, [
                "SLOT_KEY"
            ])

            handover_lab = (
                get_value_by_possible_keys(handover, [
                    "7. Phòng Lab bạn đã sử dụng là:",
                    "9. Phòng Lab bạn đã sử dụng là:",
                    "Phòng",
                    "Lab"
                ])
                or extract_lab_from_slot_key(slot_key_handover)
                or extract_lab_from_id_key(id_key)
            )

            handover_shift = (
                extract_shift_from_slot_key(slot_key_handover)
                or extract_shift_from_id_key(id_key)
            )

            report_data.append({
                "id_key": id_key,
                "lab": lab or handover_lab,
                "shift": shift or handover_shift,
                "booking_date": booking_date,
                "handover_date": get_value_by_possible_keys(handover, [
                    "Dấu thời gian",
                    "UPDATED_AT"
                ]),
                "handover_student": (
                    get_value_by_possible_keys(handover, [
                        "3. Họ và tên của bạn là:",
                        "5. Họ và tên của bạn là:",
                        "Họ và tên"
                    ])
                    or "--"
                ),
                "status": status,
                "before_image": get_value_by_possible_keys(handover, [
                    "12. Hình ảnh minh chứng trạng thái phòng Lab trước khi sử dụng:",
                    "14. Hình ảnh minh chứng trạng thái phòng Lab trước khi sử dụng:"
                ]),
                "after_image": get_value_by_possible_keys(handover, [
                    "14. Hình ảnh minh chứng trạng thái phòng Lab sau khi sử dụng/ vệ sinh:",
                    "17. Hình ảnh minh chứng trạng thái phòng Lab sau khi sử dụng/ vệ sinh:"
                ]),
                "folder_link": get_value_by_possible_keys(handover, [
                    "DRIVE_FOLDER_LINK"
                ]),
                "slot_folder_link": get_value_by_possible_keys(handover, [
                    "SLOT_FOLDER_LINK"
                ]),
                "error_message": get_value_by_possible_keys(handover, [
                    "ERROR_MESSAGE"
                ])
            })

        else:
            report_data.append({
                "id_key": id_key,
                "lab": lab,
                "shift": shift,
                "booking_date": booking_date,
                "handover_date": "--",
                "handover_student": "--",
                "status": "Chưa bàn giao",
                "before_image": "",
                "after_image": "",
                "folder_link": "",
                "slot_folder_link": "",
                "error_message": ""
            })

    return report_data


@app.route("/api/cleaning-report")
def api_cleaning_report():
    try:
        report_data = build_cleaning_records()

        return jsonify({
            "success": True,
            "data": report_data
        })

    except Exception as e:
        print("API CLEANING REPORT ERROR:", e)

        return jsonify({
            "success": False,
            "error": str(e)
        }), 500


# =========================
# API HIỆU SUẤT PHÒNG LAB
# =========================

@app.route("/api/lab-performance")
def api_lab_performance():
    try:
        bookings = get_lab_booking_rows()
        handovers = get_lab_handover_rows()

        handover_map = {}

        for handover in handovers:
            id_key = get_value_by_possible_keys(handover, [
                "1. ID_KEY của bạn là:",
                "ID_KEY",
                "ID_KEY của bạn là:"
            ])

            if id_key:
                handover_map[id_key] = handover

        performance_data = []

        for booking in bookings:
            id_key = get_value_by_possible_keys(booking, [
                "ID_KEY",
                "1. ID_KEY của bạn là:"
            ])

            if not id_key:
                continue

            status = get_value_by_possible_keys(booking, [
                "TRẠNG THÁI"
            ])

            if status != "ĐÃ_XÁC_NHẬN":
                continue

            slot_key = get_value_by_possible_keys(booking, [
                "SLOT_KEY"
            ])

            lab = (
                get_value_by_possible_keys(booking, [
                    "6. Phòng lab được lựa chọn là:",
                    "6. Phòng Lab được lựa chọn là:",
                    "Phòng",
                    "Lab"
                ])
                or extract_lab_from_slot_key(slot_key)
                or extract_lab_from_id_key(id_key)
            )

            shift = (
                get_value_by_possible_keys(booking, [
                    "9. Ca bạn đăng ký sử dụng là: ",
                    "9. Ca bạn đăng ký sử dụng là:",
                    "Ca"
                ])
                or extract_shift_from_slot_key(slot_key)
                or extract_shift_from_id_key(id_key)
            )

            booking_date = get_value_by_possible_keys(booking, [
                "8. Ngày bạn đăng ký sử dụng phòng là:",
                "Ngày đăng ký",
                "Dấu thời gian"
            ])

            purpose = get_value_by_possible_keys(booking, [
                "6. Mục đích đăng ký sử dụng phòng Lab là:",
                "6. Mục đích đăng ký sử dụng phòng lab là:",
                "Mục đích đăng ký sử dụng phòng Lab",
                "Mục đích đăng ký sử dụng phòng lab",
                "Mục đích sử dụng phòng Lab",
                "Mục đích sử dụng phòng lab",
                "Mục đích sử dụng",
                "Mục đích"
            ])

            need_support = get_value_by_possible_keys(booking, [
                "10. Bạn có cần hỗ trợ kỹ thuật không?",
                "11. Bạn có cần hỗ trợ kỹ thuật không?",
                "Cần hỗ trợ kỹ thuật",
                "Hỗ trợ kỹ thuật"
            ])

            handover = handover_map.get(id_key)

            is_handover = False
            abnormal_after = False

            if handover:
                process_status = get_value_by_possible_keys(handover, [
                    "PROCESS_STATUS"
                ])

                if process_status == "DONE":
                    is_handover = True

                abnormal_after_text = get_value_by_possible_keys(handover, [
                    "10. Phòng Lab có phát sinh vấn đề bất thường trong quá trình sử dụng không?",
                    "15. Phòng Lab có gì bất thường sau khi sử dụng không?",
                    "15. Phòng lab có gì bất thường sau khi sử dụng không?"
                ])

                abnormal_after = normalize_yes_no_py(abnormal_after_text) == "CO"

            performance_data.append({
                "id_key": id_key,
                "lab": lab,
                "shift": shift,
                "date": booking_date,
                "purpose": purpose or "Khác",
                "need_support": normalize_yes_no_py(need_support) == "CO",
                "is_handover": is_handover,
                "abnormal_after": abnormal_after
            })

        return jsonify({
            "success": True,
            "data": performance_data
        })

    except Exception as e:
        print("API LAB PERFORMANCE ERROR:", e)

        return jsonify({
            "success": False,
            "error": str(e)
        }), 500


# =========================
# GIẢNG VIÊN: DANH SÁCH LỚP + ĐỐI CHIẾU ĐIỂM DANH
# Giảng viên tải file danh sách lớp (Excel/CSV) -> lưu vào sheet LECTURER_ROSTER.
# Khi xem báo cáo, hệ thống đối chiếu MSSV trong lớp với dữ liệu đăng ký phòng Lab
# và trạng thái điểm danh admin đã cập nhật (build_attendance_records).
# =========================

ROSTER_HEADER_ALIASES = {
    "student_id": [
        "mssv", "masv", "msv", "masinhvien", "masosinhvien",
        "masosv", "studentid", "studentcode"
    ],
    "name": [
        "hoten", "hovaten", "hotensinhvien", "tensinhvien",
        "hovatensinhvien", "fullname", "name"
    ],
    "last_name": ["holot", "ho", "hodem"],
    "first_name": ["ten", "firstname"],
    "class_code": [
        "malop", "malophocphan", "malhp", "lop", "lophocphan",
        "classcode", "class", "nhomlop"
    ],
    "subject": [
        "monhoc", "tenmonhoc", "hocphan", "tenhocphan", "mon", "subject"
    ],
    "group": [
        "nhom", "tennhom", "manhom", "sttnhom", "nhomsv",
        "nhomsinhvien", "nhomthuchanh", "group", "team"
    ],
    "topic": [
        "detai", "tendetai", "detaidoan", "duan", "tenduan",
        "topic", "project"
    ]
}


def normalize_group_name(value):
    """'1' -> 'Nhóm 1'; 'nhóm 01' -> 'Nhóm 1'; giữ nguyên tên nhóm tự đặt."""
    text = " ".join(clean_value(value).split())

    if not text:
        return ""

    if text.replace(".0", "").isdigit():
        return f"Nhóm {int(float(text))}"

    key = normalize_text(text)

    if key.startswith("nhom") and key[4:].isdigit():
        return f"Nhóm {int(key[4:])}"

    return text


def normalize_student_id(value):
    """Chuẩn hóa MSSV để so khớp: bỏ khoảng trắng, bỏ đuôi .0 của Excel, viết hoa."""
    if value is None:
        return ""

    if isinstance(value, float) and value.is_integer():
        value = int(value)

    text = clean_value(value).replace(" ", "").upper()

    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]

    return text


def owner_role_label(owner_email):
    """Người tải danh sách là Giảng viên hay Kỹ thuật viên (vd. Quốc Bảo là kỹ thuật viên)."""
    lecturers = {clean_value(email).lower() for email in LECTURER_ACCOUNTS}
    return "Giảng viên" if clean_value(owner_email).lower() in lecturers else "Kỹ thuật viên"


def require_teacher_api():
    """Giảng viên và kỹ thuật viên đều dùng được API lớp học."""
    if "user" not in session:
        return jsonify({
            "success": False,
            "error": "Phiên đăng nhập đã hết. Vui lòng đăng nhập lại."
        }), 401

    return None


def current_owner_email():
    return clean_value(session.get("email") or TECH_EMAIL).lower()


def is_admin_session():
    return session.get("role", "admin") == "admin"


def get_or_create_roster_sheet():
    spreadsheet = get_booking_spreadsheet_for_write()

    try:
        sheet = spreadsheet.worksheet(ROSTER_SHEET_NAME)
    except Exception:
        sheet = spreadsheet.add_worksheet(
            title=ROSTER_SHEET_NAME,
            rows=2000,
            cols=len(ROSTER_HEADERS)
        )
        sheet.append_row(ROSTER_HEADERS)

    return sheet


def read_roster_values(sheet, fresh=False):
    # Khi chuẩn bị ghi lại cả sheet thì đọc mới (fresh) để không ghi đè dữ liệu người khác vừa sửa
    values = sheet.get_all_values() if fresh else sheet_values_cached(sheet, ROSTER_SHEET_NAME)

    if not values:
        return [ROSTER_HEADERS], []

    return [values[0]], values[1:]


def roster_rows_as_dicts(rows):
    data = []

    for row in rows:
        row = list(row) + [""] * (len(ROSTER_HEADERS) - len(row))

        data.append({
            "uploaded_at": clean_value(row[0]),
            "owner_email": clean_value(row[1]).lower(),
            "owner_name": clean_value(row[2]),
            "class_code": clean_value(row[3]),
            "subject": clean_value(row[4]),
            "student_id": normalize_student_id(row[5]),
            "name": clean_value(row[6]),
            "group": clean_value(row[7]),
            "topic": clean_value(row[8])
        })

    return data


def find_roster_columns(header_row):
    columns = {}

    for index, cell in enumerate(header_row):
        key = normalize_text(cell)

        if not key:
            continue

        for field, aliases in ROSTER_HEADER_ALIASES.items():
            if field in columns:
                continue

            if key in aliases:
                columns[field] = index
                break

            # Cho phép tiêu đề dài, ví dụ "Mã số sinh viên (MSSV)"
            if field in ("student_id", "class_code", "subject", "group", "topic") and any(
                alias in key for alias in aliases if len(alias) >= 4
            ):
                columns[field] = index
                break

    return columns


def read_uploaded_table(file_storage):
    filename = clean_value(file_storage.filename).lower()
    raw = file_storage.read()

    if filename.endswith(".xlsx") or filename.endswith(".xlsm"):
        try:
            from openpyxl import load_workbook
        except ImportError:
            raise ValueError(
                "Máy chủ chưa cài openpyxl. Chạy: py -m pip install openpyxl"
            )

        workbook = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        sheet = workbook.worksheets[0]

        return [
            list(row)
            for row in sheet.iter_rows(values_only=True)
        ]

    if filename.endswith(".csv"):
        for encoding in ("utf-8-sig", "cp1258", "latin-1"):
            try:
                text = raw.decode(encoding)
                break
            except UnicodeDecodeError:
                continue

        dialect = csv.Sniffer().sniff(text[:2048], delimiters=",;\t") if text.strip() else csv.excel
        return [row for row in csv.reader(io.StringIO(text), dialect)]

    raise ValueError("Chỉ hỗ trợ file Excel (.xlsx) hoặc CSV (.csv).")


def parse_roster_file(file_storage, default_class_code, default_subject):
    table = read_uploaded_table(file_storage)

    header_index = None
    columns = {}

    # Tìm dòng tiêu đề có cột MSSV trong 20 dòng đầu
    for index, row in enumerate(table[:20]):
        found = find_roster_columns(row)

        if "student_id" in found:
            header_index = index
            columns = found
            break

    if header_index is None:
        raise ValueError(
            "Không tìm thấy cột MSSV. Dòng tiêu đề cần có cột 'MSSV' "
            "(hoặc 'Mã sinh viên')."
        )

    if "class_code" not in columns and not default_class_code:
        raise ValueError(
            "File không có cột 'Mã lớp'. Vui lòng nhập Mã lớp ở ô phía trên."
        )

    def cell(row, field):
        index = columns.get(field)

        if index is None or index >= len(row):
            return ""

        return row[index]

    students = []
    seen = set()
    invalid_ids = []   # MSSV Văn Lang phải đúng 13 chữ số

    for offset, row in enumerate(table[header_index + 1:]):
        if not row or not any(clean_value(value) for value in row):
            continue

        student_id = normalize_student_id(cell(row, "student_id"))

        if not student_id:
            continue

        if not re.fullmatch(r"\d{13}", student_id):
            invalid_ids.append(f"dòng {header_index + offset + 2}: {student_id}")
            continue

        name = clean_value(cell(row, "name"))

        if not name and ("last_name" in columns or "first_name" in columns):
            name = " ".join(
                part for part in [
                    clean_value(cell(row, "last_name")),
                    clean_value(cell(row, "first_name"))
                ] if part
            )

        class_code = clean_value(cell(row, "class_code")) or default_class_code
        subject = clean_value(cell(row, "subject")) or default_subject
        group = normalize_group_name(cell(row, "group"))
        topic = " ".join(clean_value(cell(row, "topic")).split())

        key = (class_code.upper(), student_id)

        if key in seen:
            continue

        seen.add(key)

        students.append({
            "class_code": class_code,
            "subject": subject,
            "student_id": student_id,
            "name": name,
            "group": group,
            "topic": topic
        })

    if invalid_ids:
        more = f" (và {len(invalid_ids) - 10} dòng khác)" if len(invalid_ids) > 10 else ""
        raise ValueError(
            "MSSV phải gồm đúng 13 chữ số. Vui lòng sửa trong file rồi tải lên lại – "
            + "; ".join(invalid_ids[:10]) + more + "."
        )

    if not students:
        raise ValueError("File không có dòng sinh viên nào có MSSV.")

    if len(students) > 3000:
        raise ValueError("File quá lớn (hơn 3000 sinh viên). Vui lòng tách theo lớp.")

    return students


@app.route("/api/teacher/upload", methods=["POST"])
def teacher_upload():
    denied = require_teacher_api()
    if denied:
        return denied

    try:
        file_storage = request.files.get("file")

        if not file_storage or not file_storage.filename:
            return jsonify({
                "success": False,
                "error": "Vui lòng chọn file danh sách lớp."
            }), 400

        default_class_code = clean_value(request.form.get("class_code"))
        default_subject = clean_value(request.form.get("subject"))

        students = parse_roster_file(
            file_storage,
            default_class_code,
            default_subject
        )

        owner_email = current_owner_email()
        owner_name = clean_value(session.get("user"))
        now = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        uploaded_classes = sorted({
            item["class_code"].upper()
            for item in students
        })

        sheet = get_or_create_roster_sheet()
        header, rows = read_roster_values(sheet, fresh=True)

        # Tải lại cùng mã lớp -> thay danh sách cũ của giảng viên đó
        kept_rows = [
            row for row in rows
            if not (
                len(row) > 3
                and clean_value(row[1]).lower() == owner_email
                and clean_value(row[3]).upper() in uploaded_classes
            )
        ]

        new_rows = [
            [
                now,
                owner_email,
                owner_name,
                item["class_code"],
                item["subject"],
                item["student_id"],
                item["name"],
                item["group"],
                item["topic"]
            ]
            for item in students
        ]

        sheet.clear()
        sheet.update(
            range_name="A1",
            values=[ROSTER_HEADERS] + kept_rows + new_rows,
            value_input_option="RAW"
        )
        invalidate_cache(ROSTER_SHEET_NAME)

        group_count = len({
            (item["class_code"].upper(), item["group"])
            for item in students
            if item["group"]
        })

        class_names = sorted({item["class_code"] for item in students})

        return jsonify({
            "success": True,
            "classes": uploaded_classes,
            "student_count": len(students),
            "group_count": group_count,
            "message": (
                f"Đã lưu {len(students)} sinh viên"
                + (f", {group_count} nhóm" if group_count else "")
                + f" cho lớp {', '.join(class_names)}."
            )
        })

    except ValueError as e:
        return jsonify({
            "success": False,
            "error": str(e)
        }), 400

    except Exception as e:
        print("API TEACHER UPLOAD ERROR:", e)

        return jsonify({
            "success": False,
            "error": str(e)
        }), 500


@app.route("/api/teacher/classes")
def teacher_classes():
    denied = require_teacher_api()
    if denied:
        return denied

    try:
        sheet = get_or_create_roster_sheet()
        _, rows = read_roster_values(sheet)
        roster = roster_rows_as_dicts(rows)

        owner_email = current_owner_email()
        classes = {}

        for item in roster:
            if not is_admin_session() and item["owner_email"] != owner_email:
                continue

            key = (item["owner_email"], item["class_code"].upper())

            if key not in classes:
                classes[key] = {
                    "class_code": item["class_code"],
                    "subject": item["subject"],
                    "owner_email": item["owner_email"],
                    "owner_name": item["owner_name"],
                    "owner_role": owner_role_label(item["owner_email"]),
                    "uploaded_at": item["uploaded_at"],
                    "student_count": 0,
                    "groups": set()
                }

            classes[key]["student_count"] += 1

            if item["group"]:
                classes[key]["groups"].add(item["group"])

        for item in classes.values():
            item["group_count"] = len(item.pop("groups"))

        data = sorted(
            classes.values(),
            key=lambda item: (item["owner_name"], item["class_code"])
        )

        return jsonify({
            "success": True,
            "is_admin": is_admin_session(),
            "data": data
        })

    except Exception as e:
        print("API TEACHER CLASSES ERROR:", e)

        return jsonify({
            "success": False,
            "data": [],
            "error": str(e)
        }), 500


@app.route("/api/teacher/classes", methods=["DELETE"])
def teacher_delete_class():
    denied = require_teacher_api()
    if denied:
        return denied

    try:
        payload = request.get_json(silent=True) or {}
        class_code = clean_value(payload.get("class_code")).upper()
        owner_email = clean_value(payload.get("owner_email")).lower()

        # Giảng viên chỉ xóa được lớp của mình
        if not is_admin_session() or not owner_email:
            owner_email = current_owner_email()

        if not class_code:
            return jsonify({
                "success": False,
                "error": "Thiếu mã lớp."
            }), 400

        sheet = get_or_create_roster_sheet()
        _, rows = read_roster_values(sheet, fresh=True)

        kept_rows = [
            row for row in rows
            if not (
                len(row) > 3
                and clean_value(row[1]).lower() == owner_email
                and clean_value(row[3]).upper() == class_code
            )
        ]

        sheet.clear()
        sheet.update(
            range_name="A1",
            values=[ROSTER_HEADERS] + kept_rows,
            value_input_option="RAW"
        )
        invalidate_cache(ROSTER_SHEET_NAME)

        return jsonify({
            "success": True,
            "removed": len(rows) - len(kept_rows)
        })

    except Exception as e:
        print("API TEACHER DELETE ERROR:", e)

        return jsonify({
            "success": False,
            "error": str(e)
        }), 500


def load_class_report(class_code, owner_email):
    """Danh sách lớp + các lượt đăng ký/điểm danh của từng sinh viên.
    Trả về None nếu không có lớp. Dùng chung cho trang web và file Excel."""
    class_code = clean_value(class_code).upper()

    sheet = get_or_create_roster_sheet()
    _, rows = read_roster_values(sheet)

    roster = [
        item for item in roster_rows_as_dicts(rows)
        if item["owner_email"] == owner_email
        and item["class_code"].upper() == class_code
    ]

    if not roster:
        return None

    # Dữ liệu đăng ký + điểm danh (cùng nguồn với trang Theo dõi ra vào)
    attendance = build_attendance_records()

    sessions_by_student = {}

    for record in attendance:
        student_id = normalize_student_id(record.get("student_id"))

        if not student_id:
            continue

        sessions_by_student.setdefault(student_id, []).append({
            "booking_id": record.get("booking_id", ""),
            "lab": record.get("lab", ""),
            "date": record.get("date", ""),
            "date_iso": format_date_for_calendar(record.get("date", "")),
            "shift": record.get("shift", ""),
            "role": record.get("role", ""),
            "register_name": record.get("name", ""),
            "attendance_status": record.get("attendance_status") or "Chưa điểm danh"
        })

    students = []

    for item in roster:
        sessions = sorted(
            sessions_by_student.get(item["student_id"], []),
            key=lambda s: (s["date_iso"], s["shift"])
        )

        students.append({
            "student_id": item["student_id"],
            "name": item["name"] or (sessions[0]["register_name"] if sessions else ""),
            "group": item["group"],
            "topic": item["topic"],
            "sessions": sessions
        })

    return {
        "class_code": roster[0]["class_code"],
        "subject": roster[0]["subject"],
        "owner_name": roster[0]["owner_name"],
        "owner_role": owner_role_label(roster[0]["owner_email"]),
        "uploaded_at": roster[0]["uploaded_at"],
        "students": students
    }


def requested_owner_email(value):
    owner_email = clean_value(value).lower()

    # Giảng viên chỉ xem được lớp của mình
    if not is_admin_session() or not owner_email:
        owner_email = current_owner_email()

    return owner_email


@app.route("/api/teacher/class-report")
def teacher_class_report():
    denied = require_teacher_api()
    if denied:
        return denied

    try:
        report = load_class_report(
            request.args.get("class_code"),
            requested_owner_email(request.args.get("owner_email"))
        )

        if not report:
            return jsonify({
                "success": False,
                "error": "Không tìm thấy danh sách lớp này."
            }), 404

        return jsonify({"success": True, **report})

    except Exception as e:
        print("API TEACHER REPORT ERROR:", e)

        return jsonify({
            "success": False,
            "error": str(e)
        }), 500


@app.route("/api/teacher/template")
def teacher_template():
    """File mẫu danh sách lớp để giảng viên điền."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "DanhSachLop"

    sheet.append(["Mã lớp", "Nhóm", "Đề tài", "MSSV", "Họ tên"])
    sheet.append(["IEM101-01", "Nhóm 1", "Hệ thống chiết rót tự động", "2274802010001", "Nguyễn Văn A"])
    sheet.append(["IEM101-01", "Nhóm 1", "Hệ thống chiết rót tự động", "2274802010002", "Trần Thị B"])
    sheet.append(["IEM101-01", "Nhóm 2", "Cánh tay robot", "2274802010003", "Lê Văn C"])

    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="D72134")

    for column, width in zip("ABCDE", (14, 10, 32, 18, 28)):
        sheet.column_dimensions[column].width = width

    # Cột MSSV dạng chữ để Excel không làm mất số 0 đầu
    for row in sheet.iter_rows(min_row=2, min_col=4, max_col=4):
        for cell in row:
            cell.number_format = "@"

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="mau_danh_sach_lop.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# =========================
# XUẤT EXCEL BÁO CÁO LỚP
# Sheet: Tổng hợp / Điểm danh (bảng x - v theo ngày) / Bàn giao / Chi tiết lượt
# =========================

NO_GROUP_NAME = "Chưa phân nhóm"

# Ngưỡng mặc định cho cột "Điều kiện quá trình" (giảng viên sửa được ngay trong Excel)
DEFAULT_MAX_ABSENT_RATE = 0.2


def group_sort_key(name):
    """Sắp xếp Nhóm 1, Nhóm 2, ..., Nhóm 10; 'Chưa phân nhóm' cuối cùng."""
    if name == NO_GROUP_NAME:
        return (1, [])

    parts = re.split(r"(\d+)", normalize_text(name) or "")
    return (0, [int(part) if part.isdigit() else part for part in parts])


def summarize_student_sessions(sessions):
    """Giống cách tính trên trang web: Có mặt / Vắng / còn lại là chờ điểm danh."""
    present = sum(1 for s in sessions if s["attendance_status"] == "Có mặt")
    absent = sum(1 for s in sessions if s["attendance_status"] == "Vắng")
    pending = len(sessions) - present - absent

    if present:
        state = "Đã đến Lab"
    elif absent:
        state = "Vắng"
    elif pending:
        state = "Chờ điểm danh"
    else:
        state = "Chưa đăng ký Lab"

    return {
        "registered": len(sessions),
        "present": present,
        "absent": absent,
        "pending": pending,
        "state": state
    }


def day_mark(sessions):
    """Ký hiệu 1 ngày: x = có mặt, v = vắng, ? = đã đăng ký nhưng chưa điểm danh."""
    statuses = [s["attendance_status"] for s in sessions]

    if "Có mặt" in statuses:
        return "x"

    if "Vắng" in statuses:
        return "v"

    return "?" if statuses else ""


def split_vietnamese_name(full_name):
    parts = clean_value(full_name).split()

    if len(parts) <= 1:
        return "", " ".join(parts)

    return " ".join(parts[:-1]), parts[-1]


@app.route("/api/teacher/export")
def teacher_export():
    denied = require_teacher_api()
    if denied:
        return denied

    try:
        from openpyxl import Workbook
        from openpyxl.formatting.rule import CellIsRule, FormulaRule
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        return jsonify({
            "success": False,
            "error": "Máy chủ chưa cài openpyxl. Chạy: py -m pip install openpyxl"
        }), 500

    try:
        report = load_class_report(
            request.args.get("class_code"),
            requested_owner_email(request.args.get("owner_email"))
        )

        if not report:
            return jsonify({
                "success": False,
                "error": "Không tìm thấy danh sách lớp này."
            }), 404

        date_from = clean_value(request.args.get("date_from"))
        date_to = clean_value(request.args.get("date_to"))
        group_filter = clean_value(request.args.get("group"))

        # ---------- Lọc theo nhóm + khoảng ngày (giống bộ lọc trên trang) ----------
        students = []

        for student in report["students"]:
            group = student["group"] or NO_GROUP_NAME

            if group_filter and group != group_filter:
                continue

            sessions = [
                s for s in student["sessions"]
                if (not date_from or (s["date_iso"] and s["date_iso"] >= date_from))
                and (not date_to or (s["date_iso"] and s["date_iso"] <= date_to))
            ]

            students.append({
                **student,
                "group": group,
                "sessions": sessions,
                "summary": summarize_student_sessions(sessions)
            })

        # Giữ thứ tự trong file giảng viên tải lên, gom theo nhóm
        students.sort(key=lambda s: group_sort_key(s["group"]))
        has_groups = any(s["group"] != NO_GROUP_NAME for s in students)

        all_sessions = [s for student in students for s in student["sessions"]]
        dates = sorted({s["date_iso"] for s in all_sessions if s["date_iso"]})
        today_iso = datetime.now().strftime("%Y-%m-%d")

        def fmt_iso(iso):
            if not iso:
                return ""
            y, m, d = iso.split("-")
            return f"{d}/{m}/{y}"

        if date_from or date_to:
            period = f"{fmt_iso(date_from) or '...'} – {fmt_iso(date_to) or '...'}"
        elif dates:
            period = f"{fmt_iso(dates[0])} – {fmt_iso(dates[-1])} (toàn bộ dữ liệu)"
        else:
            period = "Toàn bộ dữ liệu"

        # ---------- Bàn giao của các lượt đăng ký có sinh viên lớp này ----------
        handover_by_id = {
            clean_value(item.get("id_key")): item
            for item in build_cleaning_records()
            if clean_value(item.get("id_key"))
        }

        bookings = {}

        for student in students:
            for session in student["sessions"]:
                booking_id = clean_value(session["booking_id"])
                if not booking_id:
                    continue

                # 1 mã có thể gồm nhiều ca -> mỗi ca 1 dòng
                booking = bookings.setdefault((booking_id, session["date_iso"], session["shift"]), {
                    "booking_id": booking_id,
                    "date_iso": session["date_iso"],
                    "date": session["date"],
                    "shift": session["shift"],
                    "lab": session["lab"],
                    "groups": set(),
                    "members": []
                })
                booking["groups"].add(student["group"])
                booking["members"].append(student["name"] or student["student_id"])

        for booking in bookings.values():
            handover = handover_by_id.get(booking["booking_id"])

            if handover:
                booking["handover_status"] = handover.get("status") or "Chưa bàn giao"
                booking["handover_date"] = "" if handover.get("handover_date") in ("--", None) else handover.get("handover_date")
                booking["handover_student"] = "" if handover.get("handover_student") in ("--", None) else handover.get("handover_student")
                booking["note"] = handover.get("error_message") or ""
            else:
                booking["handover_status"] = "Không có dữ liệu"
                booking["handover_date"] = ""
                booking["handover_student"] = ""
                booking["note"] = "Không tìm thấy lượt này trong dữ liệu bàn giao"

        booking_list = sorted(
            bookings.values(),
            key=lambda b: (b["date_iso"] or "9999", b["shift"], b["booking_id"])
        )

        # ---------- Màu & kiểu (theo mẫu bảng điểm danh của giảng viên) ----------
        NAVY = "1F3864"
        thin = Side(style="thin", color="BFBFBF")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        head_font = Font(bold=True, color="FFFFFF")
        head_fill = PatternFill("solid", fgColor=NAVY)
        green_fill = PatternFill("solid", fgColor="D8F0D8")
        yellow_fill = PatternFill("solid", fgColor="FFF2CC")
        sum_fill = PatternFill("solid", fgColor="E2EFDA")
        total_fill = PatternFill("solid", fgColor="EDEDED")
        label_fill = PatternFill("solid", fgColor="F2F2F2")
        center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left = Alignment(horizontal="left", vertical="center", wrap_text=True)

        STATE_COLORS = {
            "Đã đến Lab": ("C6EFCE", "006100"),
            "Có mặt": ("C6EFCE", "006100"),
            "Đã bàn giao": ("C6EFCE", "006100"),
            "Vắng": ("F4CCCC", "9C0006"),
            "Cần kiểm tra": ("F4CCCC", "9C0006"),
            "Chờ điểm danh": ("FFEB9C", "9C5700"),
            "Chưa điểm danh": ("FFEB9C", "9C5700"),
            "Chưa bàn giao": ("FFEB9C", "9C5700"),
            "Chưa đăng ký Lab": ("EDEDED", "595959"),
            "Không có dữ liệu": ("EDEDED", "595959"),
        }

        def style_state(cell):
            colors = STATE_COLORS.get(cell.value)
            if colors:
                cell.fill = PatternFill("solid", fgColor=colors[0])
                cell.font = Font(bold=True, color=colors[1])
                cell.alignment = center

        def header_row(ws, row, headers, start_col=1):
            for offset, title in enumerate(headers):
                cell = ws.cell(row=row, column=start_col + offset, value=title)
                cell.font = head_font
                cell.fill = head_fill
                cell.alignment = center
                cell.border = border
            ws.row_dimensions[row].height = 32

        def body_cell(ws, row, col, value, align=None, fmt=None):
            cell = ws.cell(row=row, column=col, value=value)
            cell.border = border
            cell.alignment = align or center
            if fmt:
                cell.number_format = fmt
            return cell

        def sheet_title(ws, text, width, subtitle):
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=width)
            cell = ws.cell(row=1, column=1, value=text)
            cell.font = Font(bold=True, size=14, color="FFFFFF")
            cell.fill = head_fill
            cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            ws.row_dimensions[1].height = 30
            ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=width)
            sub = ws.cell(row=2, column=1, value=subtitle)
            sub.font = Font(italic=True, color="595959")
            sub.alignment = Alignment(indent=1)

        def section_title(ws, row, text, width):
            ws.cell(row=row, column=1, value=text).font = Font(bold=True, size=12, color=NAVY)
            for col in range(1, width + 1):
                ws.cell(row=row, column=col).border = Border(bottom=Side(style="medium", color=NAVY))
            ws.row_dimensions[row].height = 22

        def ratio(part, whole):
            return (part / whole) if whole else None

        def page_setup(ws):
            ws.page_setup.orientation = "landscape"
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            ws.sheet_properties.pageSetUpPr.fitToPage = True

        class_title = report["class_code"] if normalize_text(report["class_code"]).startswith("lop") else f"Lớp {report['class_code']}"
        scope_text = group_filter or "Tất cả nhóm"
        owner_label = report["owner_role"] if report["owner_role"] == "Giảng viên" else "Người tải danh sách"
        owner_text = report["owner_name"] or "—"
        if report["owner_role"] != "Giảng viên":
            owner_text += f" ({report['owner_role']})"

        subtitle = "  ·  ".join(part for part in [
            f"Môn học: {report['subject']}" if report["subject"] else "",
            f"{owner_label}: {owner_text}",
            f"Phạm vi: {scope_text}",
            f"Thời gian: {period}"
        ] if part)

        WEEKDAYS = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6", "Thứ 7", "Chủ nhật"]
        workbook = Workbook()

        # =====================================================
        # SHEET "ĐIỂM DANH": mỗi sinh viên 1 dòng, mỗi ngày 1 cột (x / v)
        # =====================================================
        wa = workbook.active
        wa.title = "Điểm danh"
        workbook.remove(wa)
        ws_att = workbook.create_sheet("Điểm danh")

        FIRST_DATE_COL = 6                      # A STT, B MSSV, C Họ lót, D Tên, E Nhóm
        n_dates = len(dates)
        last_date_col = FIRST_DATE_COL + n_dates - 1
        col_present = FIRST_DATE_COL + n_dates
        col_absent = col_present + 1
        col_pending = col_present + 2
        col_rate = col_present + 3
        col_condition = col_present + 4
        col_missed = col_present + 5      # Lượt chưa bàn giao (KTV xem xét điểm danh)
        width_att = col_missed

        sheet_title(ws_att, f"ĐIỂM DANH PHÒNG LAB – {class_title.upper()}", width_att, subtitle)

        # Tham số (sửa trực tiếp trong Excel, các cột công thức tự tính lại)
        ws_att.cell(row=3, column=2, value="Tổng số buổi").font = Font(bold=True)
        total_cell = ws_att.cell(row=3, column=3, value=n_dates)
        total_cell.font = Font(bold=True, color="C00000")
        total_cell.fill = yellow_fill
        total_cell.border = border
        total_cell.alignment = center
        ws_att.cell(row=4, column=2, value="Vắng tối đa").font = Font(bold=True)
        max_cell = ws_att.cell(row=4, column=3, value=DEFAULT_MAX_ABSENT_RATE)
        max_cell.number_format = "0%"
        max_cell.font = Font(bold=True, color="C00000")
        max_cell.fill = yellow_fill
        max_cell.border = border
        max_cell.alignment = center
        legend = ws_att.cell(
            row=3, column=FIRST_DATE_COL,
            value="Ký hiệu:  x = có mặt   ·   v = vắng   ·   ? = đã đăng ký, chưa điểm danh   ·   trống = không đăng ký buổi đó"
        )
        ws_att.cell(row=4, column=FIRST_DATE_COL,
                    value="Sửa 2 ô vàng bên trái (tổng số buổi của học phần, % vắng tối đa) → cột Chuyên cần và Điều kiện quá trình tự tính lại."
                    ).font = Font(italic=True, color="595959")
        legend.font = Font(italic=True, color="595959")
        TOTAL_REF = "$C$3"
        MAX_REF = "$C$4"

        HEADER_ROW = 5
        first_row = HEADER_ROW + 1

        headers = ["STT", "Mã SV", "Họ lót", "Tên", "Nhóm"]
        header_row(ws_att, HEADER_ROW, headers)

        rotate = n_dates > 16
        for offset, iso in enumerate(dates):
            col = FIRST_DATE_COL + offset
            cell = ws_att.cell(row=HEADER_ROW, column=col, value=datetime.strptime(iso, "%Y-%m-%d"))
            cell.number_format = "dd/mm/yyyy"
            cell.font = head_font
            cell.fill = head_fill
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center", text_rotation=90 if rotate else 0)
            ws_att.column_dimensions[get_column_letter(col)].width = 4.6 if rotate else 11.5

        header_row(ws_att, HEADER_ROW, ["Số buổi có mặt", "Số buổi vắng", "Chưa điểm danh",
                                        "Chuyên cần", "Điều kiện quá trình", "Lượt chưa bàn giao"],
                   start_col=col_present)
        if rotate:
            ws_att.row_dimensions[HEADER_ROW].height = 72

        for index, student in enumerate(students):
            r = first_row + index
            last_name, first_name = split_vietnamese_name(student["name"])

            values = [index + 1, student["student_id"], last_name, first_name, student["group"] if has_groups else ""]
            for col, value in enumerate(values, start=1):
                cell = body_cell(ws_att, r, col, value, left if col in (3, 5) else center, "@" if col == 2 else None)
            ws_att.cell(row=r, column=4).alignment = left

            by_date = {}
            for session in student["sessions"]:
                if session["date_iso"]:
                    by_date.setdefault(session["date_iso"], []).append(session)

            for offset, iso in enumerate(dates):
                cell = body_cell(ws_att, r, FIRST_DATE_COL + offset, day_mark(by_date.get(iso, [])))
                cell.fill = yellow_fill if iso > today_iso else green_fill

            if n_dates:
                span = f"{get_column_letter(FIRST_DATE_COL)}{r}:{get_column_letter(last_date_col)}{r}"
                f_present = f'=COUNTIF({span},"x")'
                f_absent = f'=COUNTIF({span},"v")'
                f_pending = f'=COUNTIF({span},"~?")'   # "~?" = dấu ? thật (không phải ký tự đại diện)
            else:
                f_present = f_absent = f_pending = 0

            p_ref = f"{get_column_letter(col_present)}{r}"
            a_ref = f"{get_column_letter(col_absent)}{r}"

            for col, value, fmt in [
                (col_present, f_present, None),
                (col_absent, f_absent, None),
                (col_pending, f_pending, None),
                (col_rate, f"=IF({TOTAL_REF}>0,{p_ref}/{TOTAL_REF},0)", "0.00"),
                (col_condition, f'=IF({a_ref}>{TOTAL_REF}*{MAX_REF},"Không đủ điều kiện","Đủ điều kiện")', None),
            ]:
                cell = body_cell(ws_att, r, col, value, fmt=fmt)
                cell.fill = sum_fill

            # Lượt đã qua mà nhóm / sinh viên chưa bàn giao phòng
            missed_ids = {
                clean_value(session["booking_id"])
                for session in student["sessions"]
                if clean_value(session["booking_id"])
                and session["date_iso"] and session["date_iso"] < today_iso
                and bookings.get(clean_value(session["booking_id"]), {}).get("handover_status") == "Chưa bàn giao"
            }
            missed_cell = body_cell(ws_att, r, col_missed, len(missed_ids))
            if missed_ids:
                missed_cell.fill = PatternFill("solid", fgColor="F4CCCC")
                missed_cell.font = Font(bold=True, color="C00000")

        last_row = first_row + max(len(students), 1) - 1

        # Tô màu tự động (sửa x/v trong Excel thì màu đổi theo)
        if n_dates and students:
            area = f"{get_column_letter(FIRST_DATE_COL)}{first_row}:{get_column_letter(last_date_col)}{last_row}"
            ws_att.conditional_formatting.add(area, CellIsRule(operator="equal", formula=['"v"'],
                fill=PatternFill("solid", fgColor="F4CCCC"), font=Font(bold=True, color="C00000")))
            ws_att.conditional_formatting.add(area, CellIsRule(operator="equal", formula=['"x"'],
                font=Font(bold=True, color="006100")))
            ws_att.conditional_formatting.add(area, CellIsRule(operator="equal", formula=['"?"'],
                fill=PatternFill("solid", fgColor="FFEB9C"), font=Font(bold=True, color="9C5700")))

        if students:
            cond = f"{get_column_letter(col_condition)}{first_row}:{get_column_letter(col_condition)}{last_row}"
            ws_att.conditional_formatting.add(cond, CellIsRule(operator="equal", formula=['"Không đủ điều kiện"'],
                fill=PatternFill("solid", fgColor="F4CCCC"), font=Font(bold=True, color="C00000")))
            absent_area = f"{get_column_letter(col_absent)}{first_row}:{get_column_letter(col_absent)}{last_row}"
            ws_att.conditional_formatting.add(absent_area, CellIsRule(operator="greaterThan", formula=["0"],
                font=Font(bold=True, color="C00000")))
        else:
            ws_att.cell(row=first_row, column=1, value="Không có sinh viên trong phạm vi đã chọn.").font = Font(italic=True)

        for col, width in zip("ABCDE", (6, 16, 20, 9, 11 if has_groups else 4)):
            ws_att.column_dimensions[col].width = width
        for col, width in [(col_present, 10), (col_absent, 10), (col_pending, 11), (col_rate, 11), (col_condition, 20), (col_missed, 12)]:
            ws_att.column_dimensions[get_column_letter(col)].width = width

        ws_att.freeze_panes = ws_att.cell(row=first_row, column=FIRST_DATE_COL)
        ws_att.auto_filter.ref = f"A{HEADER_ROW}:{get_column_letter(width_att)}{last_row}"
        page_setup(ws_att)

        # =====================================================
        # SHEET "BÀN GIAO": mỗi lượt đăng ký của lớp 1 dòng
        # =====================================================
        ws_hand = workbook.create_sheet("Bàn giao")
        headers_h = ["STT", "Ngày", "Thứ", "Ca", "Phòng Lab", "Mã đăng ký", "Nhóm", "Số SV của lớp",
                     "Trạng thái bàn giao", "Thời gian bàn giao", "Người bàn giao", "Ghi chú",
                     "Sinh viên của lớp trong lượt"]
        sheet_title(ws_hand, f"BÀN GIAO PHÒNG LAB – {class_title.upper()}", len(headers_h), subtitle)
        header_row(ws_hand, 4, headers_h)

        for index, booking in enumerate(booking_list, start=1):
            r = 4 + index
            date_value, weekday = booking["date"], ""
            if booking["date_iso"]:
                date_value = datetime.strptime(booking["date_iso"], "%Y-%m-%d")
                weekday = WEEKDAYS[date_value.weekday()]

            groups_text = ", ".join(sorted(booking["groups"], key=group_sort_key)) if has_groups else ""
            values = [index, date_value, weekday, booking["shift"], booking["lab"], booking["booking_id"],
                      groups_text, len(booking["members"]),
                      booking["handover_status"], booking["handover_date"], booking["handover_student"],
                      booking["note"], ", ".join(booking["members"])]

            for col, value in enumerate(values, start=1):
                body_cell(ws_hand, r, col, value, left if col in (7, 11, 12, 13) else center,
                          "dd/mm/yyyy" if col == 2 else None)
            # Danh sách tên: 1 dòng, không xuống hàng
            ws_hand.cell(row=r, column=13).alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)
            ws_hand.cell(row=r, column=13).font = Font(size=9, color="595959")
            style_state(ws_hand.cell(row=r, column=9))

        if not booking_list:
            ws_hand.cell(row=5, column=1, value="Không có lượt đăng ký nào trong phạm vi đã chọn.").font = Font(italic=True)

        for col, width in enumerate([6, 12, 9, 7, 10, 22, 12, 10, 18, 20, 22, 28, 60], start=1):
            ws_hand.column_dimensions[get_column_letter(col)].width = width

        ws_hand.freeze_panes = "B5"
        ws_hand.print_area = f"A1:L{max(4 + len(booking_list), 5)}"   # in không kèm cột danh sách tên
        ws_hand.auto_filter.ref = f"A4:{get_column_letter(len(headers_h))}{max(4 + len(booking_list), 5)}"
        page_setup(ws_hand)

        # =====================================================
        # SHEET "TỔNG HỢP" (đặt đầu tiên)
        # =====================================================
        ws = workbook.create_sheet("Tổng hợp", 0)
        ws.sheet_view.showGridLines = False
        WIDTH = 11
        sheet_title(ws, "BÁO CÁO THEO DÕI SINH VIÊN SỬ DỤNG PHÒNG LAB", WIDTH,
                    "Khoa Kỹ thuật và Quản lý Công nghiệp – Trường Đại học Văn Lang")

        info = [
            ("Lớp", report["class_code"]),
            ("Môn học", report["subject"]),
            (owner_label, owner_text),
            ("Phạm vi", scope_text),
            ("Khoảng thời gian", period),
            ("Ngày xuất báo cáo", datetime.now().strftime("%d/%m/%Y %H:%M")),
        ]

        info = [(label, value) for label, value in info if value]

        row = 4
        for label, value in info:
            a = body_cell(ws, row, 1, label, left)
            a.font = Font(bold=True)
            a.fill = label_fill
            ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=5)
            body_cell(ws, row, 2, value, left)
            for col in range(3, 6):
                ws.cell(row=row, column=col).border = border
            row += 1

        size = len(students)
        registered_students = sum(1 for s in students if s["summary"]["registered"])
        visited_students = sum(1 for s in students if s["summary"]["present"])
        total_present = sum(s["summary"]["present"] for s in students)
        total_absent = sum(s["summary"]["absent"] for s in students)
        total_pending = sum(s["summary"]["pending"] for s in students)
        checked = total_present + total_absent

        hand_count = {}
        for booking in booking_list:
            hand_count[booking["handover_status"]] = hand_count.get(booking["handover_status"], 0) + 1
        n_bookings = len(booking_list)
        n_done = hand_count.get("Đã bàn giao", 0)

        att_ref = "'Điểm danh'!"
        cond_col = get_column_letter(col_condition)

        # ---------- Điểm danh (trái) + Bàn giao (phải) ----------
        row += 1
        section_title(ws, row, "ĐIỂM DANH", 4)
        ws.cell(row=row, column=7, value="BÀN GIAO").font = Font(bold=True, size=12, color=NAVY)
        for col in range(7, 12):
            ws.cell(row=row, column=col).border = Border(bottom=Side(style="medium", color=NAVY))
        row += 1
        header_row(ws, row, ["Chỉ số", "Số lượng", "Tỷ lệ", "Ghi chú"])
        header_row(ws, row, ["Chỉ số", "", "Số lượng", "Tỷ lệ"], start_col=7)
        ws.merge_cells(start_row=row, start_column=7, end_row=row, end_column=8)
        row += 1

        not_eligible = (
            f'=COUNTIF({att_ref}{cond_col}{first_row}:{cond_col}{last_row},"Không đủ điều kiện")'
            if students else 0
        )

        attendance_rows = [
            ("Sĩ số", size, None, "Sinh viên trong danh sách"),
            ("Số buổi có lịch", n_dates, None, "Ngày có sinh viên của lớp đăng ký"),
            ("SV có đăng ký Lab", registered_students, ratio(registered_students, size), "trên sĩ số"),
            ("SV đã đến Lab", visited_students, ratio(visited_students, size), "có ít nhất 1 lượt có mặt"),
            ("SV chưa đăng ký Lab", size - registered_students, ratio(size - registered_students, size), "trên sĩ số"),
            ("Lượt có mặt", total_present, ratio(total_present, checked), "trên lượt đã điểm danh"),
            ("Lượt vắng", total_absent, ratio(total_absent, checked), "trên lượt đã điểm danh"),
            ("Lượt chờ điểm danh", total_pending, None, "chưa được kỹ thuật viên điểm danh"),
            ("SV không đủ điều kiện", not_eligible, None, "theo sheet Điểm danh (vắng > ngưỡng)"),
        ]

        handover_rows = [
            ("Tổng lượt đăng ký", n_bookings, None),
            ("Đã bàn giao", n_done, ratio(n_done, n_bookings)),
            ("Chưa bàn giao", hand_count.get("Chưa bàn giao", 0), ratio(hand_count.get("Chưa bàn giao", 0), n_bookings)),
            ("Cần kiểm tra", hand_count.get("Cần kiểm tra", 0), ratio(hand_count.get("Cần kiểm tra", 0), n_bookings)),
            ("Không có dữ liệu", hand_count.get("Không có dữ liệu", 0), ratio(hand_count.get("Không có dữ liệu", 0), n_bookings)),
        ]

        for offset, (label, value, rate, note) in enumerate(attendance_rows):
            r = row + offset
            body_cell(ws, r, 1, label, left).font = Font(bold=True)
            cell = body_cell(ws, r, 2, value)
            cell.font = Font(bold=True, size=12, color="C00000" if label == "SV không đủ điều kiện" else "000000")
            body_cell(ws, r, 3, rate, fmt="0%")
            body_cell(ws, r, 4, note, left).font = Font(color="595959")

        for offset, (label, value, rate) in enumerate(handover_rows):
            r = row + offset
            ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=8)
            body_cell(ws, r, 7, label, left).font = Font(bold=True)
            ws.cell(row=r, column=8).border = border
            body_cell(ws, r, 9, value).font = Font(bold=True, size=12)
            body_cell(ws, r, 10, rate, fmt="0%")
            colors = STATE_COLORS.get(label)
            if colors:
                ws.cell(row=r, column=7).fill = PatternFill("solid", fgColor=colors[0])

        row += len(attendance_rows)

        # ---------- Theo nhóm ----------
        if has_groups:
            row += 1
            section_title(ws, row, "TỔNG HỢP THEO NHÓM", WIDTH)
            row += 1
            header_row(ws, row, [
                "Nhóm", "Đề tài", "Sĩ số", "SV đã đến Lab", "Tỷ lệ SV đến Lab",
                "Lượt có mặt", "Lượt vắng", "Chờ điểm danh",
                "Lượt đăng ký", "Đã bàn giao", "Tỷ lệ bàn giao"
            ])
            row += 1

            groups = {}
            for student in students:
                groups.setdefault(student["group"], []).append(student)

            for name in sorted(groups, key=group_sort_key):
                members = groups[name]
                g_size = len(members)
                g_visit = sum(1 for s in members if s["summary"]["present"])
                g_bookings = [b for b in booking_list if name in b["groups"]]
                g_done = sum(1 for b in g_bookings if b["handover_status"] == "Đã bàn giao")

                values = [
                    name, next((s["topic"] for s in members if s["topic"]), ""), g_size, g_visit,
                    ratio(g_visit, g_size),
                    sum(s["summary"]["present"] for s in members),
                    sum(s["summary"]["absent"] for s in members),
                    sum(s["summary"]["pending"] for s in members),
                    len(g_bookings), g_done, ratio(g_done, len(g_bookings))
                ]

                for col, value in enumerate(values, start=1):
                    cell = body_cell(ws, row, col, value, left if col <= 2 else center,
                                     "0%" if col in (5, 11) else None)
                    if col == 1:
                        cell.font = Font(bold=True)
                ws.row_dimensions[row].height = 30
                row += 1

            totals = ["TOÀN BỘ", "", size, visited_students, ratio(visited_students, size),
                      total_present, total_absent, total_pending, n_bookings, n_done, ratio(n_done, n_bookings)]

            for col, value in enumerate(totals, start=1):
                cell = body_cell(ws, row, col, value, left if col <= 2 else center,
                                 "0%" if col in (5, 11) else None)
                cell.font = Font(bold=True)
                cell.fill = total_fill
            row += 1

        # ---------- Sinh viên cần theo dõi ----------
        follow = [s for s in students if not s["summary"]["present"]]

        row += 1
        section_title(ws, row, f"SINH VIÊN CHƯA CÓ LƯỢT CÓ MẶT ({len(follow)})", WIDTH)
        row += 1

        if follow:
            header_row(ws, row, ["Nhóm", "Họ tên", "Mã SV", "Lượt đăng ký", "Lượt vắng", "Chờ điểm danh", "Tình trạng"])
            row += 1

            for student in follow:
                s = student["summary"]
                values = [student["group"] if has_groups else "", student["name"], student["student_id"],
                          s["registered"], s["absent"], s["pending"], s["state"]]

                for col, value in enumerate(values, start=1):
                    body_cell(ws, row, col, value, left if col <= 2 else center, "@" if col == 3 else None)
                style_state(ws.cell(row=row, column=7))
                row += 1
        else:
            ws.cell(row=row, column=1, value="Tất cả sinh viên đều đã có ít nhất 1 lượt có mặt.").font = Font(italic=True, color="006100")

        for col, width in enumerate([24, 40, 14, 30, 14, 11, 13, 13, 13, 13, 13], start=1):
            ws.column_dimensions[get_column_letter(col)].width = width

        page_setup(ws)

        # =====================================================
        # SHEET "CHI TIẾT LƯỢT": dữ liệu gốc từng lượt
        # =====================================================
        ws3 = workbook.create_sheet("Chi tiết lượt")
        headers3 = ["STT", "Nhóm", "Đề tài", "Mã SV", "Họ tên", "Ngày", "Thứ", "Ca",
                    "Phòng Lab", "Mã đăng ký", "Vai trò", "Điểm danh", "Bàn giao"]
        header_row(ws3, 1, headers3)
        r = 2

        for student in students:
            for session in student["sessions"]:
                date_value, weekday = session["date"], ""
                if session["date_iso"]:
                    date_value = datetime.strptime(session["date_iso"], "%Y-%m-%d")
                    weekday = WEEKDAYS[date_value.weekday()]

                booking = bookings.get(clean_value(session["booking_id"]), {})
                values = [r - 1, student["group"] if has_groups else "", student["topic"], student["student_id"],
                          student["name"], date_value, weekday, session["shift"], session["lab"],
                          session["booking_id"], session["role"], session["attendance_status"],
                          booking.get("handover_status", "")]

                for col, value in enumerate(values, start=1):
                    body_cell(ws3, r, col, value, left if col in (2, 3, 5, 11) else center,
                              "dd/mm/yyyy" if col == 6 else ("@" if col == 4 else None))
                style_state(ws3.cell(row=r, column=12))
                style_state(ws3.cell(row=r, column=13))
                r += 1

        if r == 2:
            ws3.cell(row=2, column=1, value="Không có lượt đăng ký nào trong phạm vi đã chọn.").font = Font(italic=True)

        for col, width in enumerate([6, 11, 40, 16, 26, 12, 9, 7, 10, 22, 20, 15, 15], start=1):
            ws3.column_dimensions[get_column_letter(col)].width = width

        ws3.freeze_panes = "F2"
        ws3.auto_filter.ref = f"A1:{get_column_letter(len(headers3))}{max(r - 1, 2)}"
        page_setup(ws3)

        workbook.active = 0

        # ---------- Gửi file ----------
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)

        slug = re.sub(r"[^A-Za-z0-9]+", "_", remove_vietnamese_accents(report["class_code"])).strip("_") or "lop"
        if group_filter:
            slug += "_" + re.sub(r"[^A-Za-z0-9]+", "_", remove_vietnamese_accents(group_filter)).strip("_")

        return send_file(
            output,
            as_attachment=True,
            download_name=f"bao_cao_{slug}_{datetime.now():%Y%m%d}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        print("API TEACHER EXPORT ERROR:", e)

        return jsonify({
            "success": False,
            "error": str(e)
        }), 500


# =========================
# TRANG CHỦ KỸ THUẬT VIÊN
# - Thông báo: đánh dấu "đã kiểm tra" (lưu sheet NOTIFICATION_CHECKED)
# - Yêu cầu hỗ trợ kỹ thuật: lấy từ dữ liệu đăng ký (câu hỏi "cần hỗ trợ kỹ thuật")
# - Lịch làm việc tại phòng Lab của kỹ thuật viên / trợ giảng / giảng viên
#   (sheet LAB_STAFF_SCHEDULE, kỹ thuật viên tự thêm).
#   Email: giống hệ thống Booking của sinh viên -> app chỉ ghi EMAIL_STATUS = "CHƯA GỬI",
#   Google Apps Script (file apps-script/gui-email-lich-truc.gs) gửi mail và cập nhật trạng thái.
# - Điểm danh: cột "Người điểm danh" ghi tên người trực ca đó (mặc định TECH_NAME)
# =========================

NOTIFICATION_CHECK_SHEET_NAME = "NOTIFICATION_CHECKED"
NOTIFICATION_CHECK_HEADERS = ["Mã thông báo", "Thời gian kiểm tra", "Người kiểm tra"]

STAFF_SCHEDULE_SHEET_NAME = "LAB_STAFF_SCHEDULE"
STAFF_SCHEDULE_HEADERS = [
    "Mã lịch", "Ngày", "Ca", "Phòng Lab", "Họ tên", "Vai trò",
    "Ghi chú", "Người tạo", "Thời gian tạo", "Email",
    "TRẠNG THÁI", "EMAIL_STATUS", "EMAIL_SENT_AT", "EMAIL_ERROR"
]

SCHEDULE_ACTIVE = "ĐANG HIỆU LỰC"
SCHEDULE_CANCELLED = "ĐÃ HỦY"
EMAIL_PENDING = "CHƯA GỬI"

STAFF_ROLES = ["Kỹ thuật viên", "Trợ giảng", "Giảng viên"]
STAFF_SHIFT_TIMES = {
    "Ca 1": "07:00 - 09:30",
    "Ca 2": "09:30 - 12:00",
    "Ca 3": "13:00 - 15:30",
    "Ca 4": "15:30 - 18:00"
}
EMAIL_PATTERN = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
STAFF_SHIFTS = ["Ca 1", "Ca 2", "Ca 3", "Ca 4"]
STAFF_LABS = ["Lab 1", "Lab 2", "Lab 3", "Tất cả phòng Lab"]


def booking_detail_fields(row):
    """Thông tin thêm của 1 lượt đăng ký (cùng tên cột với các API gốc)."""
    register_type = get_value_by_possible_keys(row, [
        "5. Bạn đăng ký với tư cách là:",
        "Bạn đăng ký với tư cách là:"
    ])

    student_id = get_value_by_possible_keys(row, [
        "2. Mã số sinh viên của bạn là:",
        "Mã số sinh viên của bạn là:",
        "MSSV"
    ])

    member_list = get_value_by_possible_keys(row, [
        "5.2. Danh sách thành viên nhóm là: (Họ tên 1 - MSSV 1; Họ tên 2 - MSSV 2; ...) ",
        "5.2. Danh sách thành viên nhóm là: (Họ tên 1 - MSSV 1; Họ tên 2 - MSSV 2; ...)",
        "5.2. Danh sách thành viên nhóm là:",
        "Danh sách thành viên nhóm là:"
    ])

    purpose = get_value_by_possible_keys(row, [
        "6. Mục đích đăng ký sử dụng phòng Lab là:",
        "6. Mục đích đăng ký sử dụng phòng lab là:",
        "Mục đích đăng ký sử dụng phòng Lab",
        "Mục đích đăng ký sử dụng phòng lab",
        "Mục đích sử dụng phòng Lab",
        "Mục đích sử dụng phòng lab",
        "Mục đích sử dụng",
        "Mục đích"
    ])

    support = get_value_by_possible_keys(row, [
        "10. Bạn có cần hỗ trợ kỹ thuật không?",
        "11. Bạn có cần hỗ trợ kỹ thuật không?",
        "Cần hỗ trợ kỹ thuật",
        "Hỗ trợ kỹ thuật"
    ])

    # Cách nhận biết "có cần hỗ trợ" giống API thông báo gốc
    support_text = remove_vietnamese_accents(str(support)).lower()
    need_support = "co" in support_text or "có" in str(support).lower()

    members = parse_group_members(member_list) if register_type == "Nhóm" else []

    return {
        "student_id": student_id,
        "register_type": register_type or "Cá nhân",
        "member_count": 1 + len(members),
        "purpose": purpose,
        "need_support": need_support
    }


def require_admin_api():
    if "user" not in session:
        return jsonify({
            "success": False,
            "error": "Phiên đăng nhập đã hết. Vui lòng đăng nhập lại."
        }), 401

    if not is_admin_session():
        return jsonify({
            "success": False,
            "error": "Chức năng này chỉ dành cho kỹ thuật viên."
        }), 403

    return None


def get_or_create_sheet(name, headers, cols=None):
    spreadsheet = get_booking_spreadsheet_for_write()

    try:
        return spreadsheet.worksheet(name)
    except Exception:
        sheet = spreadsheet.add_worksheet(title=name, rows=2000, cols=cols or len(headers))
        sheet.append_row(headers)
        return sheet


def ensure_sheet_headers(sheet, headers, current_header):
    """Sheet tạo từ bản trước có ít cột hơn -> cập nhật dòng tiêu đề (chỉ gọi khi ghi)."""
    if list(current_header[:len(headers)]) != list(headers):
        sheet.update(range_name="A1", values=[headers], value_input_option="RAW")


def read_checked_notifications():
    """{mã thông báo: {checked_at, checked_by}}"""
    sheet = get_or_create_sheet(NOTIFICATION_CHECK_SHEET_NAME, NOTIFICATION_CHECK_HEADERS)
    checked = {}

    for row in sheet_values_cached(sheet, NOTIFICATION_CHECK_SHEET_NAME)[1:]:
        row = list(row) + ["", "", ""]
        key = clean_value(row[0])

        if key:
            checked[key] = {
                "checked_at": clean_value(row[1]),
                "checked_by": clean_value(row[2])
            }

    return checked


@app.route("/home")
def home():
    if "user" not in session:
        return redirect(url_for("login"))

    if session.get("role") == "lecturer":
        return redirect(url_for("teacher"))

    return render_template("home.html", user=session["user"])


@app.route("/api/home/notifications")
def home_notifications():
    denied = require_admin_api()
    if denied:
        return denied

    try:
        notifications = build_notifications()
        checked = read_checked_notifications()

        for item in notifications:
            info = checked.get(item.get("notification_id"))
            item["checked"] = bool(info)
            item["checked_at"] = info["checked_at"] if info else ""
            item["checked_by"] = info["checked_by"] if info else ""

        return jsonify({"success": True, "data": notifications})

    except Exception as e:
        print("API HOME NOTIFICATIONS ERROR:", e)
        return jsonify({"success": False, "data": [], "error": str(e)}), 500


@app.route("/api/home/check", methods=["POST"])
def home_check():
    """Đánh dấu đã kiểm tra (thông báo hoặc yêu cầu hỗ trợ)."""
    denied = require_admin_api()
    if denied:
        return denied

    try:
        payload = request.get_json(silent=True) or {}
        ids = [clean_value(item) for item in payload.get("ids", []) if clean_value(item)]

        if not ids:
            return jsonify({"success": False, "error": "Thiếu mã cần đánh dấu."}), 400

        checked = read_checked_notifications()
        new_ids = [item for item in dict.fromkeys(ids) if item not in checked]

        if new_ids:
            sheet = get_or_create_sheet(NOTIFICATION_CHECK_SHEET_NAME, NOTIFICATION_CHECK_HEADERS)
            now = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            sheet.append_rows(
                [[item, now, session.get("user", "")] for item in new_ids],
                value_input_option="RAW"
            )
            invalidate_cache(NOTIFICATION_CHECK_SHEET_NAME)

        return jsonify({"success": True, "checked": len(new_ids)})

    except Exception as e:
        print("API HOME CHECK ERROR:", e)
        return jsonify({"success": False, "error": str(e)}), 500


@app.route("/api/home/support-requests")
def home_support_requests():
    denied = require_admin_api()
    if denied:
        return denied

    try:
        rows = get_lab_booking_rows()
        checked = read_checked_notifications()
        data = []

        for row in rows:
            status = get_value_by_possible_keys(row, ["TRẠNG THÁI", "TRANG THAI"])

            if status != "ĐÃ_XÁC_NHẬN":
                continue

            details = booking_detail_fields(row)

            if not details["need_support"]:
                continue

            id_key = get_value_by_possible_keys(row, ["ID_KEY", "1. ID_KEY của bạn là:"])
            slot_key = get_value_by_possible_keys(row, ["SLOT_KEY"])

            raw_date = get_value_by_possible_keys(row, [
                "8. Ngày bạn đăng ký sử dụng phòng là:",
                "Ngày bạn đăng ký sử dụng phòng là:",
                "Ngày đăng ký"
            ])

            date_iso = format_date_for_calendar(raw_date)

            if not id_key or not date_iso:
                continue

            lab = (
                get_value_by_possible_keys(row, [
                    "6. Phòng lab được lựa chọn là:",
                    "6. Phòng Lab được lựa chọn là:",
                    "Phòng lab được lựa chọn là:",
                    "Phòng Lab",
                    "Lab"
                ])
                or extract_lab_from_slot_key(slot_key)
                or extract_lab_from_id_key(id_key)
            )

            shift = (
                get_value_by_possible_keys(row, [
                    "9. Ca bạn đăng ký sử dụng là: ",
                    "9. Ca bạn đăng ký sử dụng là:",
                    "Ca bạn đăng ký sử dụng là:",
                    "Ca"
                ])
                or extract_shift_from_slot_key(slot_key)
                or extract_shift_from_id_key(id_key)
            )

            name = get_value_by_possible_keys(row, [
                "1. Họ và tên của bạn là: ",
                "1. Họ và tên của bạn là:",
                "Họ và tên của bạn là:",
                "Họ và tên"
            ])

            check_id = f"support-{id_key}"
            info = checked.get(check_id)

            data.append({
                "check_id": check_id,
                "id_key": id_key,
                "name": name,
                "lab": lab,
                "shift": shift,
                "date": raw_date,
                "date_iso": date_iso,
                "student_id": details["student_id"],
                "register_type": details["register_type"],
                "member_count": details["member_count"],
                "purpose": details["purpose"],
                "checked": bool(info),
                "checked_at": info["checked_at"] if info else "",
                "checked_by": info["checked_by"] if info else ""
            })

        data = merge_multi_shift(data, "check_id", "shift")
        data.sort(key=lambda item: (item["date_iso"], item["shift"], item["lab"]))

        return jsonify({"success": True, "data": data})

    except Exception as e:
        print("API HOME SUPPORT ERROR:", e)
        return jsonify({"success": False, "data": [], "error": str(e)}), 500


def read_staff_schedule(sheet, fresh=False, include_cancelled=False):
    values = sheet.get_all_values() if fresh else sheet_values_cached(sheet, STAFF_SCHEDULE_SHEET_NAME)
    data = []

    for row in values[1:]:
        row = list(row) + [""] * len(STAFF_SCHEDULE_HEADERS)

        if not clean_value(row[0]):
            continue

        # Ca đã hủy vẫn giữ trong sheet (để Apps Script gửi email báo hủy) nhưng không hiển thị
        if clean_value(row[10]) == SCHEDULE_CANCELLED and not include_cancelled:
            continue

        data.append({
            "id": clean_value(row[0]),
            "date_iso": format_date_for_calendar(row[1]),
            "date": clean_value(row[1]),
            "shift": clean_value(row[2]),
            "lab": clean_value(row[3]),
            "name": clean_value(row[4]),
            "role": clean_value(row[5]),
            "note": clean_value(row[6]),
            "created_by": clean_value(row[7]),
            "created_at": clean_value(row[8]),
            "email": clean_value(row[9]).lower(),
            "status": clean_value(row[10]) or SCHEDULE_ACTIVE,
            "email_status": clean_value(row[11]),
            "email_sent_at": clean_value(row[12]),
            "email_error": clean_value(row[13])
        })

    return data


@app.route("/api/staff-schedule")
def staff_schedule_list():
    denied = require_admin_api()
    if denied:
        return denied

    try:
        sheet = get_or_create_sheet(STAFF_SCHEDULE_SHEET_NAME, STAFF_SCHEDULE_HEADERS)
        data = read_staff_schedule(sheet)
        data.sort(key=lambda item: (item["date_iso"], item["shift"], item["name"]))

        return jsonify({"success": True, "data": data})

    except Exception as e:
        print("API STAFF SCHEDULE ERROR:", e)
        return jsonify({"success": False, "data": [], "error": str(e)}), 500


@app.route("/api/staff-schedule", methods=["POST"])
def staff_schedule_add():
    denied = require_admin_api()
    if denied:
        return denied

    try:
        payload = request.get_json(silent=True) or {}

        name = " ".join(clean_value(payload.get("name")).split())
        email = clean_value(payload.get("email")).lower()
        role = clean_value(payload.get("role"))
        lab = clean_value(payload.get("lab"))
        note = clean_value(payload.get("note"))
        shifts = [s for s in payload.get("shifts", []) if s in STAFF_SHIFTS]
        start = format_date_for_calendar(payload.get("date"))
        repeat_until = format_date_for_calendar(payload.get("repeat_until")) if payload.get("repeat_until") else ""

        if not name:
            return jsonify({"success": False, "error": "Vui lòng nhập họ tên."}), 400
        if email and not EMAIL_PATTERN.match(email):
            return jsonify({"success": False, "error": "Email không hợp lệ."}), 400
        if role not in STAFF_ROLES:
            return jsonify({"success": False, "error": "Vai trò không hợp lệ."}), 400
        if lab not in STAFF_LABS:
            return jsonify({"success": False, "error": "Phòng Lab không hợp lệ."}), 400
        if not start:
            return jsonify({"success": False, "error": "Vui lòng chọn ngày."}), 400
        if not shifts:
            return jsonify({"success": False, "error": "Vui lòng chọn ít nhất 1 ca."}), 400

        start_date = datetime.strptime(start, "%Y-%m-%d")
        end_date = datetime.strptime(repeat_until, "%Y-%m-%d") if repeat_until else start_date

        if end_date < start_date:
            return jsonify({"success": False, "error": "Ngày kết thúc lặp lại phải sau ngày bắt đầu."}), 400

        # Lặp lại hằng tuần, tối đa 1 học kỳ (26 tuần)
        dates = []
        current = start_date
        while current <= end_date and len(dates) < 26:
            dates.append(current)
            current += timedelta(days=7)

        sheet = get_or_create_sheet(STAFF_SCHEDULE_SHEET_NAME, STAFF_SCHEDULE_HEADERS)
        current_values = sheet.get_all_values()
        existing = {
            (item["date_iso"], item["shift"], item["name"].lower())
            for item in read_staff_schedule(sheet, fresh=True)
        }

        now = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        new_rows = []
        skipped = 0

        for date in dates:
            for shift in shifts:
                key = (date.strftime("%Y-%m-%d"), shift, name.lower())

                if key in existing:
                    skipped += 1
                    continue

                existing.add(key)
                new_rows.append([
                    uuid.uuid4().hex[:10],
                    date.strftime("%d/%m/%Y"),
                    shift,
                    lab,
                    name,
                    role,
                    note,
                    session.get("user", ""),
                    now,
                    email,
                    SCHEDULE_ACTIVE,
                    EMAIL_PENDING if email else "",
                    "",
                    ""
                ])

        if new_rows:
            ensure_sheet_headers(sheet, STAFF_SCHEDULE_HEADERS, current_values[0] if current_values else [])
            sheet.append_rows(new_rows, value_input_option="RAW")
            invalidate_cache(STAFF_SCHEDULE_SHEET_NAME)

        if new_rows:
            message = f"Đã thêm {len(new_rows)} ca làm việc cho {name}."
            if skipped:
                message += f" Bỏ qua {skipped} ca đã có lịch."
        else:
            message = f"Không thêm ca mới: {name} đã có lịch ở tất cả {skipped} ca đã chọn."

        if new_rows and email:
            message += f" Email phân công sẽ được gửi tới {email} trong ít phút."
        elif new_rows:
            message += " Không gửi email vì chưa nhập email người trực."

        return jsonify({
            "success": True,
            "added": len(new_rows),
            "skipped": skipped,
            "email_queued": bool(new_rows and email),
            "message": message
        })

    except Exception as e:
        print("API STAFF SCHEDULE ADD ERROR:", e)
        return jsonify({"success": False, "error": str(e)}), 500


@app.route("/api/staff-schedule", methods=["DELETE"])
def staff_schedule_delete():
    denied = require_admin_api()
    if denied:
        return denied

    try:
        payload = request.get_json(silent=True) or {}
        schedule_id = clean_value(payload.get("id"))

        if not schedule_id:
            return jsonify({"success": False, "error": "Thiếu mã lịch."}), 400

        sheet = get_or_create_sheet(STAFF_SCHEDULE_SHEET_NAME, STAFF_SCHEDULE_HEADERS)
        values = sheet.get_all_values()
        rows = [list(row) + [""] * (len(STAFF_SCHEDULE_HEADERS) - len(row)) for row in values[1:]]
        target = next(
            (row for row in rows if clean_value(row[0]) == schedule_id and clean_value(row[10]) != SCHEDULE_CANCELLED),
            None
        )

        if not target:
            return jsonify({"success": False, "error": "Không tìm thấy lịch này."}), 404

        # Không xóa dòng: đánh dấu ĐÃ HỦY để còn lịch sử và để Apps Script gửi email báo hủy
        # Email phân công chưa kịp gửi -> người trực chưa biết lịch, nên cũng không cần email báo hủy
        assignment_sent = clean_value(target[11]) not in ("", EMAIL_PENDING)
        target[10] = SCHEDULE_CANCELLED
        has_email = bool(clean_value(target[9])) and assignment_sent
        target[11] = EMAIL_PENDING if has_email else ""
        target[12] = ""
        target[13] = ""

        sheet.clear()
        sheet.update(range_name="A1", values=[STAFF_SCHEDULE_HEADERS] + rows, value_input_option="RAW")
        invalidate_cache(STAFF_SCHEDULE_SHEET_NAME)

        message = "Đã hủy ca làm việc."
        if has_email:
            message += f" Email báo hủy sẽ được gửi tới {clean_value(target[9])} trong ít phút."

        return jsonify({"success": True, "message": message})

    except Exception as e:
        print("API STAFF SCHEDULE DELETE ERROR:", e)
        return jsonify({"success": False, "error": str(e)}), 500


def staff_on_duty_name(raw_date, shift, lab):
    """Tên người trực ca đó theo lịch làm việc (ưu tiên đúng phòng Lab, sau đó 'Tất cả phòng Lab').
    Không có lịch trực -> mặc định kỹ thuật viên TECH_NAME."""
    try:
        date_iso = format_date_for_calendar(raw_date)
        shift_key = normalize_text(shift)
        lab_key = normalize_text(lab)

        if not date_iso or not shift_key:
            return TECH_NAME

        sheet = get_or_create_sheet(STAFF_SCHEDULE_SHEET_NAME, STAFF_SCHEDULE_HEADERS)
        on_shift = [
            item for item in read_staff_schedule(sheet)
            if item["date_iso"] == date_iso and normalize_text(item["shift"]) == shift_key
        ]

        same_lab = [item for item in on_shift if normalize_text(item["lab"]) == lab_key]
        all_labs = [item for item in on_shift if item["lab"] == "Tất cả phòng Lab"]
        people = same_lab or all_labs

        names = list(dict.fromkeys(item["name"] for item in people if item["name"]))
        return ", ".join(names) if names else TECH_NAME

    except Exception as e:
        print("STAFF ON DUTY ERROR:", e)
        return TECH_NAME


# =========================
# QUY ĐỊNH BÀN GIAO: GHI CHÚ ĐỎ CHO KỸ THUẬT VIÊN KHI ĐIỂM DANH
# - Lượt đã qua (trước hôm nay) mà chưa có phiếu bàn giao -> "chưa bàn giao".
# - Đếm số lượt chưa bàn giao của từng sinh viên trong tháng của buổi đó
#   (lượt nhóm tính cho tất cả thành viên).
# - Lần 1, 2: KTV xem xét điểm danh; đủ MISSED_HANDOVER_LIMIT lần: cổng sinh viên
#   tạm khóa đăng ký đến hết tháng.
# =========================

MISSED_HANDOVER_LIMIT = 3


def attach_handover_flags(records):
    status_by_id = {
        clean_value(item.get("id_key")): item.get("status") or ""
        for item in build_cleaning_records()
        if clean_value(item.get("id_key"))
    }

    today = datetime.now().date()

    # Cổng sinh viên: nhóm chọn "từng thành viên tự bàn giao" -> xét theo từng MSSV
    modes = portal_handover_modes()
    handed_pairs = portal_handover_pairs() if modes else set()

    def student_key(record):
        return normalize_student_id(record.get("student_id")) or normalize_text(record.get("name"))

    def valid_date(record):
        # Ngày sinh viên nhập sai năm (vd 07/08/0002, 21/02/2005) -> không tính,
        # chỉ báo KTV kiểm tra lại dữ liệu. Dữ liệu làm mới 4 tháng/lần nên ±1 năm là đủ.
        parsed = parse_notification_date(record.get("date"))
        if parsed and abs(parsed.year - today.year) <= 1:
            return parsed
        return None

    def month_of(record):
        parsed = valid_date(record)
        return parsed.strftime("%m/%Y") if parsed else ""

    def is_missed(record):
        booking_id = clean_value(record.get("booking_id"))
        parsed = valid_date(record)
        if not (booking_id and parsed and parsed.date() < today):
            return False

        if modes.get(booking_id) == "TỪNG THÀNH VIÊN":
            return (booking_id, normalize_student_id(record.get("student_id"))) not in handed_pairs

        return status_by_id.get(booking_id) == "Chưa bàn giao"

    missed_bookings = {}

    for record in records:
        if is_missed(record):
            key = (student_key(record), month_of(record))
            missed_bookings.setdefault(key, set()).add(clean_value(record.get("booking_id")))

    for record in records:
        month = month_of(record)
        count = len(missed_bookings.get((student_key(record), month), set()))
        record["date_invalid"] = valid_date(record) is None
        record["handover_missed"] = is_missed(record)
        record["missed_in_month"] = count if month else 0
        record["missed_month"] = month
        record["missed_limit"] = MISSED_HANDOVER_LIMIT
        record["handover_mode"] = modes.get(clean_value(record.get("booking_id")), "")

    return records


# =========================
# ĐỌC THÊM DỮ LIỆU TỪ CỔNG SINH VIÊN (đăng ký trên web)
# - Cổng sinh viên ghi vào CHÍNH file Booking của kỹ thuật viên, ở các sheet riêng:
#   WEB_BOOKING (mỗi ca 1 dòng), WEB_MEMBERS (từng thành viên), WEB_HANDOVER.
#   Sheet LAB_BOOKING / LAB_HANDOVER của Google Form giữ nguyên.
# - Hệ thống chuyển các dòng đó về đúng tên cột của form cũ rồi gộp chung,
#   nên mọi trang (Lịch, Theo dõi ra vào, Vệ sinh & bàn giao, Hiệu suất,
#   Trang chủ, Giảng viên) dùng được ngay mà không đổi cách tính gốc.
# - Chưa có sheet WEB_... (chưa ai đăng ký trên web) thì bỏ qua, không báo lỗi.
# - Muốn tắt: để PORTAL_SPREADSHEET_ID = "".
# =========================

PORTAL_SPREADSHEET_ID = BOOKING_SPREADSHEET_ID
PORTAL_BOOKING_SHEET = "WEB_BOOKING"
PORTAL_MEMBER_SHEET = "WEB_MEMBERS"
PORTAL_HANDOVER_SHEET = "WEB_HANDOVER"
PORTAL_SOURCE = "STUDENT_PORTAL"
PORTAL_SOURCE_LABEL = "Cổng sinh viên"
PORTAL_RETRY_SECONDS = 120

_portal_state = {"failed_at": 0.0, "error": ""}
_portal_missing = {}

FORM_MEMBER_LIST_KEY = "5.2. Danh sách thành viên nhóm là: (Họ tên 1 - MSSV 1; Họ tên 2 - MSSV 2; ...)"
FORM_ABNORMAL_KEY = "10. Phòng Lab có phát sinh vấn đề bất thường trong quá trình sử dụng không?"


def _portal_rows(sheet_name):
    """Đọc 1 sheet của cổng sinh viên (có bộ nhớ đệm). Lỗi -> [] và tạm nghỉ 2 phút."""
    if not PORTAL_SPREADSHEET_ID:
        return []

    if time.time() - _portal_state["failed_at"] < PORTAL_RETRY_SECONDS:
        return []

    if time.time() - _portal_missing.get(sheet_name, 0) < 60:
        return []

    try:
        return get_sheet_rows(PORTAL_SPREADSHEET_ID, sheet_name)
    except Exception as e:
        if "WorksheetNotFound" in type(e).__name__ or "WorksheetNotFound" in str(e):
            # Chưa có sheet (chưa ai đăng ký / bàn giao trên web): nhớ 60 giây để đỡ tốn lượt đọc
            _portal_missing[sheet_name] = time.time()
            return []
        _portal_state["failed_at"] = time.time()
        _portal_state["error"] = str(e)
        print("PORTAL SHEET ERROR:", sheet_name, e)
        return []


def _portal_date(value):
    """Ngày của cổng (2026-10-29 hoặc 29/10/2026) -> dd/mm/yyyy như form cũ."""
    text = clean_value(value)
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S", "%m/%d/%Y"):
        try:
            return datetime.strptime(text, fmt).strftime("%d/%m/%Y")
        except Exception:
            pass
    return text


def _portal_datetime(value):
    text = clean_value(value)
    for fmt in ("%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt).strftime("%d/%m/%Y %H:%M:%S")
        except Exception:
            pass
    return text


def _portal_web_bookings():
    return [
        row for row in _portal_rows(PORTAL_BOOKING_SHEET)
        if clean_value(row.get("SOURCE")) == PORTAL_SOURCE and clean_value(row.get("ID_KEY"))
    ]


def portal_booking_rows(existing_ids=None):
    """Lượt đăng ký trên cổng sinh viên, theo tên cột của Google Form cũ."""
    bookings = _portal_web_bookings()
    if not bookings:
        return []

    existing_ids = existing_ids or set()

    members_by_id = {}
    for member in _portal_rows(PORTAL_MEMBER_SHEET):
        members_by_id.setdefault(clean_value(member.get("ID_KEY")), []).append(member)

    rows = []

    for row in bookings:
        id_key = clean_value(row.get("ID_KEY"))
        if id_key in existing_ids:
            continue

        rep_id = clean_value(row.get("MSSV_DAI_DIEN"))
        is_group = remove_vietnamese_accents(clean_value(row.get("HINH_THUC"))).lower() == "nhom"

        others = [
            f"{clean_value(m.get('MEMBER_NAME'))} - {clean_value(m.get('MEMBER_MSSV'))}"
            for m in members_by_id.get(id_key, [])
            if clean_value(m.get("MEMBER_MSSV")) and clean_value(m.get("MEMBER_MSSV")) != rep_id
        ]

        status = clean_value(row.get("BOOKING_STATUS")).upper().replace(" ", "_")

        rows.append({
            "Dấu thời gian": _portal_datetime(row.get("CREATED_AT")),
            "TRẠNG THÁI": status,
            "ID_KEY": id_key,
            "SLOT_KEY": clean_value(row.get("SLOT_KEY")),
            "1. Họ và tên của bạn là:": clean_value(row.get("TEN_DAI_DIEN")),
            "2. Mã số sinh viên của bạn là:": rep_id,
            "3. Email trường cấp cho bạn là:": clean_value(row.get("EMAIL_TRUONG")) or clean_value(row.get("EMAIL_NHAN_XAC_NHAN")),
            "4. Số điện thoại của bạn là:": clean_value(row.get("SDT_DAI_DIEN")),
            "5. Bạn đăng ký với tư cách là:": "Nhóm" if is_group else "Cá nhân",
            "5.1. Số lượng thành viên trong nhóm là:": clean_value(row.get("MEMBER_COUNT")),
            FORM_MEMBER_LIST_KEY: "; ".join(others) if is_group else "",
            "6. Mục đích đăng ký sử dụng phòng Lab là:": clean_value(row.get("MUC_DICH")),
            "6. Phòng lab được lựa chọn là:": clean_value(row.get("PHONG")),
            "7. Khu vực bạn đăng ký sử dụng trên phòng Lab là:": clean_value(row.get("KHU_VUC")),
            "8. Ngày bạn đăng ký sử dụng phòng là:": _portal_date(row.get("NGAY_SU_DUNG")),
            "9. Ca bạn đăng ký sử dụng là:": clean_value(row.get("CA_SU_DUNG")),
            "10. Bạn có cần hỗ trợ kỹ thuật không?": clean_value(row.get("HO_TRO_KY_THUAT")),
            "HINH_THUC_BAN_GIAO": clean_value(row.get("HINH_THUC_BAN_GIAO")),
            "MA_DANG_KY_CHUNG": clean_value(row.get("MA_DANG_KY_CHUNG")),
            "NGUỒN": PORTAL_SOURCE_LABEL
        })

    return rows


def _portal_batches():
    """ID_KEY -> mã chung của các ca đăng ký cùng lúc (không có thì chính nó)."""
    batch_of = {}
    for row in _portal_web_bookings():
        id_key = clean_value(row.get("ID_KEY"))
        batch_of[id_key] = clean_value(row.get("MA_DANG_KY_CHUNG")) or id_key
    return batch_of


def portal_handover_rows(existing_ids=None):
    """Phiếu bàn giao trên cổng sinh viên, theo tên cột của form bàn giao cũ.
    Các ca đăng ký cùng lúc bàn giao 1 lần -> phiếu được áp cho mọi ca trong lô."""
    handovers = [
        row for row in _portal_rows(PORTAL_HANDOVER_SHEET)
        if clean_value(row.get("ID_KEY"))
    ]
    if not handovers:
        return []

    existing_ids = existing_ids or set()
    batch_of = _portal_batches()
    ids_in_batch = {}
    for id_key, batch in batch_of.items():
        ids_in_batch.setdefault(batch, []).append(id_key)

    rows = []

    for row in handovers:
        source_id = clean_value(row.get("ID_KEY"))
        condition = clean_value(row.get("TINH_TRANG"))
        note = clean_value(row.get("GHI_CHU"))
        cond_text = remove_vietnamese_accents(condition).lower()
        problem = "su co" in cond_text
        clean_ok = cond_text.startswith("sach")

        batch = batch_of.get(source_id, source_id)

        for id_key in ids_in_batch.get(batch, [source_id]):
            if id_key in existing_ids:
                continue

            rows.append({
                "Dấu thời gian": _portal_datetime(row.get("CREATED_AT")),
                "1. ID_KEY của bạn là:": id_key,
                "3. Họ và tên của bạn là:": clean_value(row.get("TEN_SINH_VIEN")),
                "MSSV": clean_value(row.get("MSSV")),
                "7. Phòng Lab bạn đã sử dụng là:": clean_value(row.get("PHONG")),
                "PROCESS_STATUS": "DONE" if clean_ok else "CHECK",
                "ERROR_MESSAGE": (f"{condition}: {note}" if note else condition) if problem else "",
                FORM_ABNORMAL_KEY: "Có" if problem else "Không",
                "TINH_TRANG": condition,
                "GHI_CHU": note,
                "MA_BAN_GIAO": clean_value(row.get("HANDOVER_ID")),
                "BAN_GIAO_CHUNG_VOI": source_id if id_key != source_id else "",
                "NGUỒN": PORTAL_SOURCE_LABEL
            })

    return rows


def portal_handover_modes():
    """ID_KEY -> hình thức bàn giao (ĐẠI DIỆN / TỪNG THÀNH VIÊN / CÁ NHÂN)."""
    return {
        clean_value(row.get("ID_KEY")): clean_value(row.get("HINH_THUC_BAN_GIAO")).upper()
        for row in _portal_web_bookings()
    }


def portal_handover_pairs():
    """(ID_KEY, MSSV) đã bàn giao – áp cho mọi ca trong cùng lô."""
    batch_of = _portal_batches()
    ids_in_batch = {}
    for id_key, batch in batch_of.items():
        ids_in_batch.setdefault(batch, []).append(id_key)

    pairs = set()
    for row in _portal_rows(PORTAL_HANDOVER_SHEET):
        source_id = clean_value(row.get("ID_KEY"))
        mssv = normalize_student_id(row.get("MSSV"))
        for id_key in ids_in_batch.get(batch_of.get(source_id, source_id), [source_id]):
            pairs.add((id_key, mssv))
    return pairs


def merge_multi_shift(items, key_field, shift_field=None):
    """Cổng sinh viên: 1 mã đăng ký có thể gồm nhiều ca (mỗi ca 1 dòng, cùng ID_KEY).
    Gộp các mục trùng mã thành 1 mục, ghép tên ca: "Ca 2, Ca 3, Ca 4"."""
    first_by_key = {}
    result = []

    for item in items:
        key = item.get(key_field)

        if not key or key not in first_by_key:
            if key:
                first_by_key[key] = item
            result.append(item)
            continue

        first = first_by_key[key]

        if shift_field:
            shift = clean_value(item.get(shift_field))
            current = clean_value(first.get(shift_field))
            if shift and shift not in current.split(", "):
                first[shift_field] = f"{current}, {shift}" if current else shift
        else:
            head, sep, shift = clean_value(item.get("message")).rpartition(" - ")
            first_head, first_sep, first_shift = clean_value(first.get("message")).rpartition(" - ")
            if sep and first_sep and shift and shift not in first_shift.split(", "):
                first["message"] = f"{first_head} - {first_shift}, {shift}"

    return result


def portal_source_status():
    return {
        "enabled": bool(PORTAL_SPREADSHEET_ID),
        "spreadsheet_id": PORTAL_SPREADSHEET_ID,
        "error": _portal_state["error"] if time.time() - _portal_state["failed_at"] < PORTAL_RETRY_SECONDS else ""
    }


# =========================
# API TEST KẾT NỐI SHEET
# =========================

@app.route("/api/test-sheets")
def test_sheets():
    try:
        booking_rows = get_lab_booking_rows()
        handover_rows = get_lab_handover_rows()

        return jsonify({
            "success": True,
            "booking_sheet_id": BOOKING_SPREADSHEET_ID,
            "handover_sheet_id": HANDOVER_SPREADSHEET_ID,
            "booking_sheet_name": BOOKING_SHEET_NAME,
            "handover_sheet_name": HANDOVER_SHEET_NAME,
            "booking_rows": len(booking_rows),
            "handover_rows": len(handover_rows),
            "portal_source": portal_source_status()
        })

    except Exception as e:
        return jsonify({
            "success": False,
            "error": str(e)
        }), 500


# =========================
# RUN
# =========================

if __name__ == "__main__":
    app.run(debug=True)