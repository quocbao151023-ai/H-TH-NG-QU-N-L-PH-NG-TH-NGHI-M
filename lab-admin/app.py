from flask import Flask, render_template, jsonify, request, redirect, url_for, session, send_file
from datetime import datetime, timedelta
import csv
import io
import os
import re
import threading
import time
import uuid
import unicodedata
import gspread
from google.oauth2.service_account import Credentials

# Giờ Việt Nam cho mọi datetime.now() (Render chạy giờ UTC) – thời gian ghi vào sheet,
# "hôm nay", "qua ngày"… khớp với cổng sinh viên (VIETNAM_TZ).
os.environ.setdefault("TZ", "Asia/Ho_Chi_Minh")
if hasattr(time, "tzset"):
    time.tzset()

app = Flask(__name__)
# Nên đặt SECRET_KEY / TECH_PASSWORD trong biến môi trường (Render: Environment)
app.secret_key = os.environ.get("SECRET_KEY", "lab-system-secret-key")

# Giới hạn file tải lên (danh sách lớp) tối đa 5 MB
app.config["MAX_CONTENT_LENGTH"] = 5 * 1024 * 1024

# =========================
# TÀI KHOẢN ĐĂNG NHẬP
# =========================

TECH_EMAIL = "Quocbao151023@gmail.com"
TECH_PASSWORD = os.environ.get("TECH_PASSWORD", "p4ssw0rd")
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

    if not os.path.exists(GOOGLE_CREDENTIALS_FILE):
        raise RuntimeError(
            f"Không tìm thấy file {GOOGLE_CREDENTIALS_FILE} trong thư mục chạy app.py. "
            "Chép file khóa Google (key.json) vào cùng thư mục với app.py rồi chạy lại."
        )

    credentials = Credentials.from_service_account_file(
        GOOGLE_CREDENTIALS_FILE,
        scopes=scopes
    )

    return gspread.authorize(credentials)


# Kết nối ghi được tạo ở lần dùng đầu (thiếu key.json thì app vẫn mở được và báo lỗi rõ ràng)
_write_client = None


def get_write_client():
    global _write_client

    if _write_client is None:
        _write_client = get_google_client(readonly=False)

    return _write_client


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
        spreadsheet = CachedSpreadsheet(get_write_client().open_by_key(BOOKING_SPREADSHEET_ID))
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
                "lab": lab,
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
                "lab": lab,
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
                "lab": (
                    get_value_by_possible_keys(row, [
                        "7. Phòng Lab bạn đã sử dụng là:",
                        "9. Phòng Lab bạn đã sử dụng là:",
                        "Phòng",
                        "Lab"
                    ])
                    or extract_lab_from_id_key(id_key)
                ),
                "level": "danger"
            })

    # Vật tư BÁO KTV cần thu hồi (trang Mượn – trả)
    try:
        notifications.extend(loan_recall_notifications())
    except Exception as e:
        print("LOAN RECALL NOTIFICATION ERROR:", e)

    # Phiếu bàn giao chờ KTV kiểm tra 5S / sinh viên đã làm lại
    try:
        notifications.extend(ktv_check_notifications())
    except Exception as e:
        print("KTV CHECK NOTIFICATION ERROR:", e)

    # Hộp / rổ sinh viên báo sắp hết (trang Vật tư)
    try:
        notifications.extend(low_stock_notifications())
    except Exception as e:
        print("LOW STOCK NOTIFICATION ERROR:", e)

    # Đánh giá 5S định kỳ: đến hạn chấm, việc khắc phục quá hạn, lỗi lặp lại khi bàn giao (SOP-05)
    try:
        notifications.extend(audit_notifications())
    except Exception as e:
        print("AUDIT NOTIFICATION ERROR:", e)

    # Phạm vi phòng: mặc định chỉ Lab 1 (nút "Xem thêm Lab 2, Lab 3" mở lại)
    notifications = [item for item in notifications if in_lab_scope(item.get("lab"))]

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

    # Sinh viên KTV bổ sung vào lượt (nhóm đăng ký thiếu thành viên)
    data.extend(extra_member_records(data, latest_status))

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
    ktv_checks = handover_ktv_checks()

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
                ]),
                # Sinh viên tự khai (phiếu bàn giao theo SOP-02)
                "handover_id": clean_value(handover.get("MA_BAN_GIAO")),
                "self_condition": clean_value(handover.get("TINH_TRANG")),
                "responsible_area": clean_value(handover.get("KHU_VUC_PHU_TRACH")),
                "zones": clean_value(handover.get("KHU_KIEM_TRA")),
                "checklist": clean_value(handover.get("BANG_KIEM_DAT")),
                "error_codes": clean_value(handover.get("MA_LOI")),
                "failed_items": clean_value(handover.get("MUC_CHUA_DAT")),
                "checklist_detail": clean_value(handover.get("BANG_KIEM_CHI_TIET")),
                "low_stock": clean_value(handover.get("VI_TRI_SAP_HET")),
                "open_loans": clean_value(handover.get("MON_CHUA_TRA")),
                "note": clean_value(handover.get("GHI_CHU")),
                # Sinh viên báo đã làm lại sau khi KTV ghi Chưa đạt
                "redo_at": clean_value(handover.get("SV_LAM_LAI_LUC")),
                "redo_note": clean_value(handover.get("SV_LAM_LAI_GHI_CHU")),
                "redo_count": clean_value(handover.get("SO_LAN_LAM_LAI")),
                "from_portal": clean_value(handover.get("NGUỒN")) == PORTAL_SOURCE_LABEL,
                # KTV xác nhận – lưu riêng
                **ktv_check_fields(ktv_checks.get(id_key))
            })
            # Vòng KTV kiểm tra 5S chỉ áp dụng phiếu từ cổng sinh viên (có bảng kiểm);
            # phiếu Google Form cũ vẫn xác nhận được nếu KTV muốn, nhưng không tính là "đang chờ".
            report_data[-1]["ktv_state"] = (
                ktv_state_of(report_data[-1])
                if report_data[-1]["from_portal"] or report_data[-1]["ktv_result"] else ""
            )

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
                "error_message": "",
                "ktv_state": "",
                **ktv_check_fields(None)
            })

    return report_data


# Vòng kiểm tra 5S sau bàn giao (đồng bộ với cổng sinh viên):
# Chờ KTV kiểm tra -> Đạt | Chưa đạt (SV làm lại) -> SV báo đã làm lại -> Chờ kiểm tra lại -> …
KTV_STATE_WAITING = "cho_ktv"
KTV_STATE_REDO = "lam_lai"
KTV_STATE_RECHECK = "cho_kiem_lai"
KTV_STATE_PASSED = "dat"
KTV_STATE_LABELS = {
    KTV_STATE_WAITING: "Chờ KTV kiểm tra",
    KTV_STATE_REDO: "Chưa đạt – chờ SV làm lại",
    KTV_STATE_RECHECK: "SV đã làm lại – kiểm tra lại",
    KTV_STATE_PASSED: "Đạt",
}
KTV_NOTIFY_DAYS = 3   # chỉ nhắc các phiếu bàn giao trong 3 ngày gần nhất


def ktv_state_of(record):
    result = record.get("ktv_result")
    if result == KTV_PASS:
        return KTV_STATE_PASSED
    if result == KTV_FAIL:
        redo_at = parse_datetime_any(record.get("redo_at"))
        ktv_at = parse_datetime_any(record.get("ktv_at"))
        return KTV_STATE_RECHECK if redo_at and ktv_at and redo_at > ktv_at else KTV_STATE_REDO
    return KTV_STATE_WAITING


def ktv_check_notifications():
    """Phiếu bàn giao mới cần KTV kiểm tra 5S (chuông + Trang chủ)."""
    since = datetime.now() - timedelta(days=KTV_NOTIFY_DAYS)
    items = []

    for record in build_cleaning_records():
        state = record.get("ktv_state")
        if state not in (KTV_STATE_WAITING, KTV_STATE_RECHECK):
            continue
        # Chỉ phiếu từ cổng sinh viên (có bảng kiểm) – phiếu Google Form cũ không nằm trong vòng này
        if not record.get("from_portal"):
            continue
        handed = parse_datetime_any(record.get("handover_date"))
        if not handed or handed < since:
            continue

        recheck = state == KTV_STATE_RECHECK
        items.append({
            "notification_id": f"ktv-{record['id_key']}-{record.get('redo_at') if recheck else record.get('handover_date')}",
            "type": "ktv_recheck" if recheck else "ktv_check",
            "title": "SV ĐÃ LÀM LẠI – KIỂM TRA LẠI 5S" if recheck else "CHỜ KTV KIỂM TRA 5S",
            "message": (
                f"{record['lab']} · {record['shift']} · {record.get('handover_student') or ''}: "
                + (f"báo đã làm lại {record.get('ktv_codes') or ''}".strip() if recheck
                   else f"bảng kiểm {record.get('checklist') or '--'}"
                   + (f", tự khai lỗi {record['error_codes']}" if record.get("error_codes") else ""))
                + ". Sinh viên đang chờ kết quả để ra về."
            ),
            "id_key": record["id_key"],
            "date": record.get("redo_at") if recheck else record.get("handover_date"),
            "lab": record["lab"],
            "level": "danger" if recheck else "warning",
            "link": url_for("cleaning_report", q=record["id_key"])
        })

    return items


def handover_ktv_checks():
    """ID_KEY -> dòng KTV xác nhận gần nhất (sheet BAN_GIAO_KTV). Lỗi đọc -> {} (không chặn trang)."""
    try:
        _, rows, _ = read_table(HANDOVER_CHECK_SHEET, HANDOVER_CHECK_HEADERS)
    except Exception as e:
        print("HANDOVER KTV CHECK READ ERROR:", e)
        return {}

    return {row["ID_KEY"]: row for row in rows}


def ktv_check_fields(row):
    row = row or {}
    return {
        "ktv_result": row.get("Kết quả KTV", ""),
        "ktv_codes": row.get("Mã lỗi KTV", ""),
        "ktv_note": row.get("Ghi chú", ""),
        "ktv_by": row.get("KTV", ""),
        "ktv_at": row.get("Thời gian", "")
    }


@app.route("/api/handover/ktv-check", methods=["POST"])
def api_handover_ktv_check():
    """KTV xác nhận Đạt / Không đạt 5S cho 1 lượt bàn giao (đối chiếu ảnh và hiện trạng).
    Lưu riêng ở sheet BAN_GIAO_KTV, không sửa trạng thái sinh viên tự khai."""
    denied = require_admin_api()
    if denied:
        return denied

    try:
        payload = request.get_json(silent=True) or {}
        id_key = clean_value(payload.get("id_key"))
        result = clean_value(payload.get("result"))
        codes = ", ".join(
            code for code in dict.fromkeys(
                re.sub(r"\s+", "", part).upper()
                for part in clean_value(payload.get("codes")).replace(";", ",").split(",")
            ) if code
        )
        note = clean_value(payload.get("note"))

        if not id_key:
            return json_error("Thiếu mã đăng ký.")
        if result not in (KTV_PASS, KTV_FAIL, ""):
            return json_error("Kết quả phải là Đạt hoặc Không đạt.")
        if result == KTV_FAIL and not (codes or note):
            return json_error("Không đạt: ghi mã lỗi (vd B-S2) hoặc ghi chú lý do.")

        sheet, rows, header = read_table(HANDOVER_CHECK_SHEET, HANDOVER_CHECK_HEADERS, fresh=True)
        record = next((row for row in rows if row["ID_KEY"] == id_key), None)
        values = {
            "ID_KEY": id_key,
            "Mã bàn giao": clean_value(payload.get("handover_id")),
            "Kết quả KTV": result,
            "Mã lỗi KTV": codes if result == KTV_FAIL else "",
            "Ghi chú": note,
            "KTV": session.get("user", ""),
            "Thời gian": now_text()
        }

        if record:
            record.update(values)
            update_record(sheet, HANDOVER_CHECK_SHEET, HANDOVER_CHECK_HEADERS, header, record)
        else:
            append_records(sheet, HANDOVER_CHECK_SHEET, HANDOVER_CHECK_HEADERS, header, [values])

        return jsonify({"success": True, **ktv_check_fields(values)})  # trạng thái vòng kiểm tra tính lại ở trang

    except Exception as e:
        print("API HANDOVER KTV CHECK ERROR:", e)
        return json_error(str(e), 500)


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
    "Ca 4": "15:30 - 17:00"
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

            if not in_lab_scope(lab):
                continue

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

# Cột phiếu bàn giao theo SOP Lab 1 – 5S (cổng sinh viên ghi thêm vào WEB_HANDOVER)
PORTAL_5S_FIELDS = [
    "KHU_VUC_PHU_TRACH", "KHU_KIEM_TRA", "BANG_KIEM_DAT", "MA_LOI", "MUC_CHUA_DAT",
    "BANG_KIEM_CHI_TIET", "VI_TRI_SAP_HET", "SU_CO_VI_TRI", "SU_CO_LOAI",
    "SU_CO_DA_DUNG_MAY", "MON_CHUA_TRA",
    "SV_LAM_LAI_LUC", "SV_LAM_LAI_GHI_CHU", "SO_LAN_LAM_LAI"
]


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

        # Chỉ sự cố mới là "vấn đề bàn giao"; mục 5S chưa đạt (mã B-S2…) đi theo vòng KTV kiểm tra 5S
        error_message = (f"{condition}: {note}" if note else condition) if problem else ""

        extra = {key: clean_value(row.get(key)) for key in PORTAL_5S_FIELDS}

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
                "ERROR_MESSAGE": error_message,
                FORM_ABNORMAL_KEY: "Có" if problem else "Không",
                "TINH_TRANG": condition,
                "GHI_CHU": note,
                "MA_BAN_GIAO": clean_value(row.get("HANDOVER_ID")),
                "BAN_GIAO_CHUNG_VOI": source_id if id_key != source_id else "",
                "NGUỒN": PORTAL_SOURCE_LABEL,
                **extra
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
# PHẠM VI PHÒNG
# - Mặc định chỉ hiện Lab 1 (phòng triển khai 5S) trên hệ thống quản lý.
# - Nút "Xem thêm Lab 2, Lab 3" ghi cookie lab_scope=all (theo từng trình duyệt).
# - Chỉ lọc khi HIỂN THỊ; dữ liệu gốc của 3 phòng trong Google Sheet giữ nguyên.
# =========================

DEFAULT_LAB = "Lab 1"
LAB_SCOPE_COOKIE = "lab_scope"


def lab_scope_all():
    try:
        return request.cookies.get(LAB_SCOPE_COOKIE) == "all"
    except RuntimeError:
        # Gọi ngoài request (vd. chạy thử) -> không lọc
        return True


def normalize_lab_name(value):
    text = normalize_text(value)

    for number in ("1", "2", "3"):
        if f"lab{number}" in text:
            return f"Lab {number}"

    return clean_value(value)


def in_lab_scope(lab):
    if lab_scope_all():
        return True

    lab = normalize_lab_name(lab)
    return not lab or lab == DEFAULT_LAB or normalize_text(lab) == normalize_text("Tất cả phòng Lab")


# =========================
# MÁY KHU C – LAB 1 (decal đỏ C1 – C7)
# =========================

MACHINES = [
    {"code": "C1", "name": "Máy in 3D"},
    {"code": "C2", "name": "Phay CNC mini"},
    {"code": "C3", "name": "Tiện mini"},
    {"code": "C4", "name": "Cưa lọng bàn"},
    {"code": "C5", "name": "Máy cắt góc"},
    {"code": "C6", "name": "Khoan bàn"},
    # C7: máy cắt – theo sơ đồ phân khu Lab 1 (cập nhật 11/2026)
    {"code": "C7", "name": "Máy cắt"},
]
MACHINE_BY_CODE = {item["code"]: item for item in MACHINES}

AREA_TARGET = {"code": "PHONG", "name": "Phòng / khu vực chung"}


# =========================
# VẬT TƯ – THIẾT BỊ, MƯỢN – TRẢ, NHẬT KÝ SỰ CỐ
# Nguyên tắc: hệ thống là nơi DUY NHẤT lưu dữ liệu vật tư.
# - Không điền sẵn số liệu: danh mục, số lượng chuẩn, tối thiểu, số đếm đều do KTV nhập.
# - Các sheet tự tạo trong file Booking ở lần dùng đầu tiên.
# =========================

ITEM_GROUP_SELF = "TỰ LẤY"     # dưới 1 triệu / món – nhãn xanh, sinh viên tự lấy, không ghi mượn
ITEM_GROUP_ASK = "BÁO KTV"     # trên 1 triệu / món – nhãn đỏ, KTV lấy giúp và ghi mượn – trả
ITEM_GROUPS = [ITEM_GROUP_SELF, ITEM_GROUP_ASK]

ITEM_ACTIVE = "ĐANG DÙNG"
ITEM_INACTIVE = "NGỪNG DÙNG"

# Quá hạn kiểm kê: không có lần kiểm kê nào trong số ngày này (kiểm kê hằng tháng)
INVENTORY_COUNT_DAYS = 31

LOAN_OPEN = "ĐANG MƯỢN"
LOAN_RETURNED = "ĐÃ TRẢ"
RETURN_CONDITIONS = ["Bình thường", "Hỏng", "Thiếu"]

# Mốc báo "cần thu hồi" cho món BÁO KTV chưa trả (đã chốt 10/11/2026: LÚC BÀN GIAO).
# Phiếu bàn giao trên cổng sinh viên cũng hiện danh sách món còn mượn để nhắc trả trước khi gửi.
# Các mốc hỗ trợ (đổi nếu cần):
#   "het_ca"   – đã hết ca của lượt đăng ký
#   "ban_giao" – lượt đã gửi phiếu bàn giao phòng mà còn món chưa trả
#   "qua_ngay" – món mượn từ hôm trước
LOAN_RECALL_RULES = ["ban_giao"]
LOAN_RECALL_TEXT = {
    "het_ca": "Đã hết ca",
    "ban_giao": "Lượt đã bàn giao phòng",
    "qua_ngay": "Mượn từ hôm trước"
}

INCIDENT_NEW = "Mới"
INCIDENT_PROGRESS = "Đang xử lý"
INCIDENT_DONE = "Đã xong"
INCIDENT_STATUSES = [INCIDENT_NEW, INCIDENT_PROGRESS, INCIDENT_DONE]

SOURCE_HANDOVER = "Phiếu bàn giao"
SOURCE_RETURN = "Trả vật tư"
SOURCE_STAFF = "KTV ghi nhận"

INVENTORY_SHEET = "VAT_TU_DANH_MUC"
INVENTORY_HEADERS = [
    "Mã vật tư", "Tên vật tư", "Nhóm", "Mã vị trí", "Đơn vị",
    "SL chuẩn", "SL tối thiểu", "Phòng Lab", "Ghi chú", "Trạng thái",
    "Người cập nhật", "Cập nhật lúc", "Đã bổ sung lúc"
]

COUNT_SHEET = "VAT_TU_KIEM_KE"
COUNT_HEADERS = [
    "Thời gian", "Kỳ", "Mã vật tư", "Tên vật tư", "SL đếm", "SL hỏng",
    "Đang cho mượn lúc kiểm", "Người kiểm", "Ghi chú"
]

LOAN_SHEET = "MUON_TRA"
LOAN_HEADERS = [
    "Mã mượn", "Thời gian mượn", "ID_KEY", "Phòng Lab", "Ngày sử dụng", "Ca",
    "MSSV", "Họ tên", "Mã vật tư", "Tên vật tư", "Số lượng", "Người ghi mượn",
    "Trạng thái", "Thời gian trả", "Tình trạng khi trả", "Người nhận trả",
    "Ghi chú", "Mã sự cố"
]

INCIDENT_SHEET = "NHAT_KY_SU_CO"
INCIDENT_HEADERS = [
    "Mã sự cố", "Thời gian ghi nhận", "Nguồn", "Mã nguồn", "Phòng Lab",
    "Đối tượng", "Tên đối tượng", "Mô tả", "Người báo", "Trạng thái",
    "Cách xử lý", "Người xử lý", "Ngày xong", "Cập nhật lúc", "Loại sự cố"
]

# Loại sự cố – SOP-04 mục 2 (dùng chung với phiếu bàn giao trên cổng sinh viên)
INCIDENT_TYPES = ["Khẩn cấp", "Hỏng thiết bị", "Bất thường nhỏ"]

# Vị trí khác máy C theo sơ đồ 5S Lab 1 (11/2026) – để ghi sự cố đúng mã
POSITION_TARGETS = (
    [{"code": f"B{i}", "name": f"Bàn B{i}"} for i in range(1, 11)]
    + [
        {"code": "T1", "name": "Tủ T1 (D1)"},
        {"code": "T2", "name": "Tủ T2 (D2)"},
        {"code": "K1", "name": "Kệ K1 (D3)"},
        {"code": "K2", "name": "Kệ K2 (D3)"},
        {"code": "K3", "name": "Kệ K3 (D3)"},
        {"code": "D4", "name": "Vật liệu tấm & dài (D4)"},
        {"code": "F1", "name": "Dụng cụ vệ sinh (F1)"},
        {"code": "F2", "name": "Máy hút bụi (F2)"},
        {"code": "A", "name": "Khu giảng viên (A)"},
        {"code": "E", "name": "Khu ghế xếp (E)"},
    ]
)
POSITION_BY_CODE = {item["code"]: item for item in POSITION_TARGETS}

# KTV xác nhận đạt 5S cho từng phiếu bàn giao – lưu RIÊNG với trạng thái sinh viên tự khai (SOP-02 mục 7)
HANDOVER_CHECK_SHEET = "BAN_GIAO_KTV"
HANDOVER_CHECK_HEADERS = [
    "ID_KEY", "Mã bàn giao", "Kết quả KTV", "Mã lỗi KTV", "Ghi chú", "KTV", "Thời gian"
]
KTV_PASS = "Đạt"
KTV_FAIL = "Không đạt"

LOCATION_PATTERN = re.compile(r"^(T|K)\d+-N\d+-\d{2}$")


def now_text():
    return datetime.now().strftime("%d/%m/%Y %H:%M:%S")


def read_table(name, headers, fresh=False):
    """Đọc sheet hệ thống -> (sheet, [dict theo tên cột + _row])."""
    sheet = get_or_create_sheet(name, headers)
    values = sheet.get_all_values() if fresh else sheet_values_cached(sheet, name)
    rows = []

    for index, row in enumerate(values[1:], start=2):
        row = list(row) + [""] * (len(headers) - len(row))

        if not clean_value(row[0]):
            continue

        record = {header: clean_value(row[i]) for i, header in enumerate(headers)}
        record["_row"] = index
        rows.append(record)

    header = values[0] if values else []
    return sheet, rows, header


def append_records(sheet, name, headers, current_header, records):
    ensure_sheet_headers(sheet, headers, current_header)
    sheet.append_rows(
        [[record.get(h, "") for h in headers] for record in records],
        value_input_option="RAW"
    )
    invalidate_cache(name)


def update_record(sheet, name, headers, current_header, record):
    ensure_sheet_headers(sheet, headers, current_header)
    sheet.update(
        range_name=f"A{record['_row']}",
        values=[[record.get(h, "") for h in headers]],
        value_input_option="RAW"
    )
    invalidate_cache(name)


def parse_quantity(value, allow_blank=False):
    """Số lượng KTV nhập: số nguyên >= 0. Trống -> None (nếu cho phép)."""
    text = clean_value(value).replace(",", ".")

    if text == "":
        if allow_blank:
            return None
        raise ValueError("Vui lòng nhập số lượng.")

    try:
        number = float(text)
    except ValueError:
        raise ValueError(f"Số lượng không hợp lệ: {text}")

    if number < 0 or not number.is_integer():
        raise ValueError(f"Số lượng phải là số nguyên không âm: {text}")

    return int(number)


def to_int(value):
    try:
        return int(float(clean_value(value)))
    except Exception:
        return None


def parse_datetime_any(value):
    text = clean_value(value)

    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt)
        except Exception:
            pass

    return None


def json_error(message, code=400):
    return jsonify({"success": False, "error": message}), code


# ---------- Danh mục + kiểm kê ----------

def low_stock_reports():
    """Mã vị trí -> các lần sinh viên báo sắp hết trên phiếu bàn giao (cột VI_TRI_SAP_HET)."""
    reports = {}
    try:
        rows = _portal_rows(PORTAL_HANDOVER_SHEET)
    except Exception as e:
        print("LOW STOCK REPORT ERROR:", e)
        return reports

    for row in rows:
        codes = clean_value(row.get("VI_TRI_SAP_HET"))
        if not codes:
            continue
        at = _portal_datetime(row.get("CREATED_AT"))
        for code in codes.replace(";", ",").split(","):
            code = code.strip().upper()
            if code:
                reports.setdefault(code, []).append({
                    "at": at,
                    "when": parse_datetime_any(at),
                    "id_key": clean_value(row.get("ID_KEY")),
                    "student": clean_value(row.get("TEN_SINH_VIEN"))
                })

    for items in reports.values():
        items.sort(key=lambda report: report["when"] or datetime.min, reverse=True)
    return reports


def low_stock_notifications():
    return [
        {
            "notification_id": f"low-stock-{item['id']}-{item['low_reports'][0]['at']}",
            "type": "low_stock",
            "title": "HỘP / RỔ SẮP HẾT",
            "message": f"{item['location']} · {item['name']}: sinh viên báo dưới vạch đỏ ({item['low_reports'][0]['student'] or item['low_reports'][0]['id_key']})",
            "id_key": item["low_reports"][0]["id_key"],
            "date": item["low_reports"][0]["at"],
            "lab": item["lab"],
            "level": "warning"
        }
        for item in build_inventory() if "sv_bao_het" in item["flags"]
    ]


def build_inventory():
    _, items, _ = read_table(INVENTORY_SHEET, INVENTORY_HEADERS)
    _, counts, _ = read_table(COUNT_SHEET, COUNT_HEADERS)
    _, loans, _ = read_table(LOAN_SHEET, LOAN_HEADERS)
    _, incidents, _ = read_table(INCIDENT_SHEET, INCIDENT_HEADERS)

    on_loan = {}
    for loan in loans:
        if loan["Trạng thái"] == LOAN_OPEN:
            on_loan[loan["Mã vật tư"]] = on_loan.get(loan["Mã vật tư"], 0) + (to_int(loan["Số lượng"]) or 0)

    last_count = {}
    for count in counts:
        when = parse_datetime_any(count["Thời gian"])
        current = last_count.get(count["Mã vật tư"])
        if when and (not current or when >= current["_when"]):
            last_count[count["Mã vật tư"]] = dict(count, _when=when)

    damage_open = {}
    for incident in incidents:
        # Món hỏng: sự cố vật tư chưa xong (bỏ qua sự cố "Thiếu khi trả" – thiếu thể hiện qua số kiểm kê)
        if (incident["Trạng thái"] != INCIDENT_DONE and incident["Đối tượng"].startswith("VT-")
                and not incident["Mô tả"].startswith("Thiếu")):
            damage_open[incident["Đối tượng"]] = damage_open.get(incident["Đối tượng"], 0) + 1

    today = datetime.now()
    data = []
    low_reports = low_stock_reports()

    for item in items:
        item_id = item["Mã vật tư"]
        standard = to_int(item["SL chuẩn"])
        minimum = to_int(item["SL tối thiểu"])
        count = last_count.get(item_id)
        loaned = on_loan.get(item_id, 0)

        flags = []
        counted = to_int(count["SL đếm"]) if count else None
        damaged = to_int(count["SL hỏng"]) if count else None
        effective = None

        if count is None or counted is None:
            flags.append("chua_kiem_ke")
        else:
            # Món đang cho mượn lúc kiểm vẫn thuộc kho -> cộng lại khi so với chuẩn
            effective = counted + (to_int(count["Đang cho mượn lúc kiểm"]) or 0)

            if (today - count["_when"]).days > INVENTORY_COUNT_DAYS:
                flags.append("qua_han")

            if minimum is not None and effective < minimum:
                flags.append("duoi_toi_thieu")
            elif standard is not None and effective < standard:
                flags.append("thieu")
            elif standard is not None:
                flags.append("du")

        if (damaged or 0) > 0 or damage_open.get(item_id):
            flags.append("co_hong")

        if loaned:
            flags.append("dang_muon")

        # Sinh viên báo hộp / rổ dưới vạch đỏ trên phiếu bàn giao (SOP-03 bước 4).
        # Hết báo khi KTV bấm "Đã bổ sung" hoặc kiểm kê món đó sau lần báo.
        handled_at = max(
            [when for when in [count["_when"] if count else None, parse_datetime_any(item["Đã bổ sung lúc"])] if when],
            default=None
        )
        reports = [
            report for report in low_reports.get(item["Mã vị trí"].upper(), [])
            if not handled_at or (report["when"] and report["when"] > handled_at)
        ]
        if reports and item["Trạng thái"] != ITEM_INACTIVE:
            flags.append("sv_bao_het")

        data.append({
            "id": item_id,
            "name": item["Tên vật tư"],
            "group": item["Nhóm"],
            "location": item["Mã vị trí"],
            "unit": item["Đơn vị"],
            "standard_qty": standard,
            "min_qty": minimum,
            "lab": item["Phòng Lab"] or DEFAULT_LAB,
            "note": item["Ghi chú"],
            "active": item["Trạng thái"] != ITEM_INACTIVE,
            "updated_by": item["Người cập nhật"],
            "updated_at": item["Cập nhật lúc"],
            "last_count": counted,
            "last_damaged": damaged,
            "last_count_at": count["Thời gian"] if count else "",
            "last_count_by": count["Người kiểm"] if count else "",
            "effective_qty": effective,
            "on_loan": loaned,
            "open_damage_incidents": damage_open.get(item_id, 0),
            "restocked_at": item["Đã bổ sung lúc"],
            "low_reports": [
                {"at": report["at"], "id_key": report["id_key"], "student": report["student"]}
                for report in reports
            ],
            "flags": flags
        })

    data.sort(key=lambda item: (not item["active"], item["location"] or "~", item["name"]))
    return data


def inventory_page_guard():
    if "user" not in session:
        return redirect(url_for("login"))

    if session.get("role") == "lecturer":
        return redirect(url_for("teacher"))

    return None


@app.route("/vat-tu")
def inventory_page():
    return inventory_page_guard() or render_template("inventory.html", user=session["user"])


@app.route("/muon-tra")
def loans_page():
    return inventory_page_guard() or render_template("loans.html", user=session["user"])


@app.route("/su-co")
def incidents_page():
    # Nhật ký sự cố đã gộp thành một thẻ trong trang Vật tư
    denied = inventory_page_guard()
    if denied:
        return denied
    query = request.args.get("q", "").strip()
    return redirect(url_for("inventory_page", tab="su-co", **({"q": query} if query else {})))


@app.route("/api/inventory")
def api_inventory():
    denied = require_admin_api()
    if denied:
        return denied

    try:
        items = [item for item in build_inventory() if in_lab_scope(item["lab"])]

        return jsonify({
            "success": True,
            "data": items,
            "groups": ITEM_GROUPS,
            "count_days": INVENTORY_COUNT_DAYS
        })

    except Exception as e:
        print("API INVENTORY ERROR:", e)
        return json_error(str(e), 500)


@app.route("/api/inventory", methods=["POST"])
def api_inventory_save():
    """Thêm / sửa 1 món trong danh mục (không có id = thêm mới)."""
    denied = require_admin_api()
    if denied:
        return denied

    try:
        payload = request.get_json(silent=True) or {}

        item_id = clean_value(payload.get("id"))
        name = " ".join(clean_value(payload.get("name")).split())
        group = clean_value(payload.get("group")).upper()
        location = clean_value(payload.get("location")).upper().replace(" ", "")
        unit = clean_value(payload.get("unit"))
        lab = normalize_lab_name(payload.get("lab")) or DEFAULT_LAB
        note = clean_value(payload.get("note"))
        active = payload.get("active", True) is not False

        if not name:
            return json_error("Vui lòng nhập tên vật tư.")
        if group not in ITEM_GROUPS:
            return json_error("Vui lòng chọn nhóm TỰ LẤY hoặc BÁO KTV.")
        if not location:
            return json_error("Vui lòng nhập mã vị trí (vd. T1-N2-04, K2-N2-03).")
        if not unit:
            return json_error("Vui lòng nhập đơn vị tính.")

        standard = parse_quantity(payload.get("standard_qty"))
        minimum = parse_quantity(payload.get("min_qty"), allow_blank=True)

        if minimum is not None and minimum > standard:
            return json_error("Số lượng tối thiểu không được lớn hơn số lượng chuẩn.")

        sheet, items, header = read_table(INVENTORY_SHEET, INVENTORY_HEADERS, fresh=True)

        duplicate = next(
            (item for item in items
             if item["Mã vật tư"] != item_id
             and normalize_text(item["Tên vật tư"]) == normalize_text(name)
             and item["Mã vị trí"] == location),
            None
        )
        if duplicate:
            return json_error(f"Đã có \"{name}\" tại {location} ({duplicate['Mã vật tư']}).")

        record = {
            "Tên vật tư": name,
            "Nhóm": group,
            "Mã vị trí": location,
            "Đơn vị": unit,
            "SL chuẩn": standard,
            "SL tối thiểu": "" if minimum is None else minimum,
            "Phòng Lab": lab,
            "Ghi chú": note,
            "Trạng thái": ITEM_ACTIVE if active else ITEM_INACTIVE,
            "Người cập nhật": session.get("user", ""),
            "Cập nhật lúc": now_text()
        }

        if item_id:
            target = next((item for item in items if item["Mã vật tư"] == item_id), None)
            if not target:
                return json_error("Không tìm thấy vật tư này.", 404)
            target.update(record)
            update_record(sheet, INVENTORY_SHEET, INVENTORY_HEADERS, header, target)
        else:
            numbers = [to_int(item["Mã vật tư"].replace("VT-", "")) or 0 for item in items]
            item_id = f"VT-{(max(numbers) if numbers else 0) + 1:04d}"
            record["Mã vật tư"] = item_id
            append_records(sheet, INVENTORY_SHEET, INVENTORY_HEADERS, header, [record])

        return jsonify({
            "success": True,
            "id": item_id,
            "location_warning": "" if LOCATION_PATTERN.match(location) else
                "Mã vị trí chưa đúng dạng Tủ/Kệ-Tầng-Số (vd. T1-N2-04, K2-N2-03)."
        })

    except ValueError as e:
        return json_error(str(e))
    except Exception as e:
        print("API INVENTORY SAVE ERROR:", e)
        return json_error(str(e), 500)


# ---------- Nhập danh mục vật tư từ file Excel / CSV ----------

INVENTORY_IMPORT_COLUMNS = [
    # (khóa, tiêu đề trong file mẫu, các tên cột chấp nhận – đã bỏ dấu, bỏ khoảng trắng)
    ("id", "Mã vật tư (để trống nếu thêm mới)", ["mavattu", "mavt", "ma"]),
    ("name", "Tên vật tư", ["tenvattu", "tenvt", "ten", "vattu"]),
    ("group", "Nhóm (TỰ LẤY / BÁO KTV)", ["nhom", "nhomvattu"]),
    ("location", "Mã vị trí (T1-N2-04 / K2-N2-03)", ["mavitri", "vitri"]),
    ("unit", "Đơn vị tính", ["donvitinh", "donvi", "dvt"]),
    ("standard_qty", "SL chuẩn", ["slchuan", "soluongchuan", "soluong", "sl"]),
    ("min_qty", "SL tối thiểu", ["sltoithieu", "soluongtoithieu", "toithieu"]),
    ("lab", "Phòng Lab", ["phonglab", "phong", "lab"]),
    ("note", "Ghi chú", ["ghichu"]),
]


def inventory_import_columns(header_row):
    """Tìm vị trí từng cột theo tiêu đề (bỏ dấu, bỏ phần trong ngoặc)."""
    keys = [normalize_text(re.sub(r"\(.*?\)", "", clean_value(cell))) for cell in header_row]
    columns = {}

    for field, _, aliases in INVENTORY_IMPORT_COLUMNS:
        for index, key in enumerate(keys):
            if key in aliases and index not in columns.values():
                columns[field] = index
                break

    return columns


def parse_inventory_import(file_storage, existing_items):
    table = [row for row in read_uploaded_table(file_storage) if any(clean_value(c) for c in (row or []))]

    header_index, columns = None, {}
    for index, row in enumerate(table[:10]):
        found = inventory_import_columns(row)
        if "name" in found and "location" in found:
            header_index, columns = index, found
            break

    if header_index is None:
        raise ValueError("Không tìm thấy dòng tiêu đề có cột \"Tên vật tư\" và \"Mã vị trí\". Hãy dùng file mẫu.")

    by_id = {item["Mã vật tư"]: item for item in existing_items}
    by_key = {(normalize_text(item["Tên vật tư"]), item["Mã vị trí"]): item for item in existing_items}
    seen = set()
    rows = []

    for offset, raw in enumerate(table[header_index + 1:], start=header_index + 2):
        cell = lambda field: clean_value(raw[columns[field]]) if field in columns and columns[field] < len(raw) else ""

        name = " ".join(cell("name").split())
        group_text = normalize_text(cell("group"))
        group = ITEM_GROUP_SELF if group_text in ("tulay", "xanh") else ITEM_GROUP_ASK if group_text in ("baoktv", "do") else cell("group").upper()
        location = cell("location").upper().replace(" ", "")
        unit = cell("unit")
        errors = []

        if not name:
            errors.append("thiếu tên vật tư")
        if group not in ITEM_GROUPS:
            errors.append("nhóm phải là TỰ LẤY hoặc BÁO KTV")
        if not location:
            errors.append("thiếu mã vị trí")
        if not unit:
            errors.append("thiếu đơn vị tính")

        standard = minimum = None
        try:
            standard = parse_quantity(cell("standard_qty"))
        except ValueError as e:
            errors.append("SL chuẩn: " + str(e).replace("Vui lòng nhập số lượng.", "chưa nhập"))
        try:
            minimum = parse_quantity(cell("min_qty"), allow_blank=True)
        except ValueError as e:
            errors.append("SL tối thiểu: " + str(e))
        if standard is not None and minimum is not None and minimum > standard:
            errors.append("SL tối thiểu lớn hơn SL chuẩn")

        key = (normalize_text(name), location)
        if name and location and key in seen:
            errors.append("trùng dòng khác trong file")
        seen.add(key)

        target = by_id.get(cell("id").upper()) or by_key.get(key)

        rows.append({
            "row": offset,
            "id": target["Mã vật tư"] if target else "",
            "action": "update" if target else "new",
            "name": name,
            "group": group,
            "location": location,
            "unit": unit,
            "standard_qty": standard,
            "min_qty": minimum,
            "lab": normalize_lab_name(cell("lab")) or DEFAULT_LAB,
            "note": cell("note"),
            "errors": errors,
            "location_warning": bool(location) and not LOCATION_PATTERN.match(location)
        })

    if not rows:
        raise ValueError("File không có dòng vật tư nào bên dưới dòng tiêu đề.")

    return rows


@app.route("/api/inventory/import", methods=["POST"])
def api_inventory_import():
    """Tải file lên. mode=preview: chỉ kiểm tra. mode=save: ghi các dòng hợp lệ vào danh mục."""
    denied = require_admin_api()
    if denied:
        return denied

    try:
        file_storage = request.files.get("file")
        if not file_storage or not file_storage.filename:
            return json_error("Vui lòng chọn file Excel (.xlsx) hoặc CSV.")

        sheet, items, header = read_table(INVENTORY_SHEET, INVENTORY_HEADERS, fresh=True)
        rows = parse_inventory_import(file_storage, items)
        valid = [row for row in rows if not row["errors"]]

        summary = {
            "total": len(rows),
            "new": sum(1 for row in valid if row["action"] == "new"),
            "update": sum(1 for row in valid if row["action"] == "update"),
            "error": len(rows) - len(valid)
        }

        if request.form.get("mode") != "save":
            return jsonify({"success": True, "rows": rows, "summary": summary})

        if not valid:
            return json_error("Không có dòng hợp lệ để nhập.")

        stamp = {"Người cập nhật": session.get("user", ""), "Cập nhật lúc": now_text()}
        numbers = [to_int(item["Mã vật tư"].replace("VT-", "")) or 0 for item in items]
        next_number = (max(numbers) if numbers else 0) + 1
        by_id = {item["Mã vật tư"]: item for item in items}
        new_records = []

        for row in valid:
            record = {
                "Tên vật tư": row["name"],
                "Nhóm": row["group"],
                "Mã vị trí": row["location"],
                "Đơn vị": row["unit"],
                "SL chuẩn": row["standard_qty"],
                "SL tối thiểu": "" if row["min_qty"] is None else row["min_qty"],
                "Phòng Lab": row["lab"],
                "Ghi chú": row["note"],
                **stamp
            }

            if row["action"] == "update":
                target = by_id[row["id"]]
                target.update(record)
                update_record(sheet, INVENTORY_SHEET, INVENTORY_HEADERS, header, target)
            else:
                record["Mã vật tư"] = f"VT-{next_number:04d}"
                record["Trạng thái"] = ITEM_ACTIVE
                next_number += 1
                new_records.append(record)

        if new_records:
            append_records(sheet, INVENTORY_SHEET, INVENTORY_HEADERS, header, new_records)

        return jsonify({"success": True, "summary": summary})

    except ValueError as e:
        return json_error(str(e))
    except Exception as e:
        print("API INVENTORY IMPORT ERROR:", e)
        return json_error(str(e), 500)


@app.route("/api/inventory/import-template")
def api_inventory_import_template():
    """File mẫu chỉ có dòng tiêu đề + trang hướng dẫn – không có số liệu mẫu."""
    denied = require_admin_api()
    if denied:
        return denied

    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation

    workbook = Workbook()
    ws = workbook.active
    ws.title = "Danh muc vat tu"
    ws.append([title for _, title, _ in INVENTORY_IMPORT_COLUMNS])

    for index, width in enumerate([22, 30, 24, 28, 14, 12, 14, 12, 30], start=1):
        cell = ws.cell(row=1, column=index)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="D72134")
        ws.column_dimensions[cell.column_letter].width = width
    ws.freeze_panes = "A2"

    group_rule = DataValidation(type="list", formula1='"TỰ LẤY,BÁO KTV"', allow_blank=True)
    lab_rule = DataValidation(type="list", formula1='"Lab 1,Lab 2,Lab 3"', allow_blank=True)
    ws.add_data_validation(group_rule)
    ws.add_data_validation(lab_rule)
    group_rule.add("C2:C1000")
    lab_rule.add("H2:H1000")

    guide = workbook.create_sheet("Huong dan")
    for line in [
        "Điền mỗi món 1 dòng ở trang \"Danh muc vat tu\". Không đổi tên cột.",
        "Mã vật tư: để trống khi thêm mới. Ghi mã VT-xxxx (hoặc trùng Tên + Mã vị trí) thì hệ thống cập nhật món đó.",
        "Nhóm: TỰ LẤY (dưới 1 triệu, nhãn xanh) hoặc BÁO KTV (trên 1 triệu, nhãn đỏ).",
        "Mã vị trí: Tủ T / Kệ K – tầng N – số hộp, rổ. Vd. T1-N2-04, K2-N2-03. Không ghi mã phòng.",
        "SL chuẩn: bắt buộc, số nguyên. SL tối thiểu: tùy chọn, không lớn hơn SL chuẩn.",
        "Phòng Lab: để trống = Lab 1.",
        "Số đếm kiểm kê không nhập ở file này – dùng nút \"Kiểm kê tháng\" trên trang Vật tư.",
    ]:
        guide.append([line])
    guide.column_dimensions["A"].width = 110

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="mau_nhap_vat_tu.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/api/inventory/restocked", methods=["POST"])
def api_inventory_restocked():
    """KTV đã bổ sung hộp / rổ sinh viên báo sắp hết -> tắt cảnh báo."""
    denied = require_admin_api()
    if denied:
        return denied

    try:
        item_id = clean_value((request.get_json(silent=True) or {}).get("id"))
        sheet, items, header = read_table(INVENTORY_SHEET, INVENTORY_HEADERS, fresh=True)
        target = next((item for item in items if item["Mã vật tư"] == item_id), None)
        if not target:
            return json_error("Không tìm thấy vật tư này.", 404)

        target["Đã bổ sung lúc"] = now_text()
        target["Người cập nhật"] = session.get("user", "")
        update_record(sheet, INVENTORY_SHEET, INVENTORY_HEADERS, header, target)
        return jsonify({"success": True})

    except Exception as e:
        print("API INVENTORY RESTOCKED ERROR:", e)
        return json_error(str(e), 500)


@app.route("/api/inventory/count", methods=["POST"])
def api_inventory_count():
    """Kiểm kê: KTV nhập số đếm thực tế. Món để trống = chưa kiểm lần này (không ghi)."""
    denied = require_admin_api()
    if denied:
        return denied

    try:
        payload = request.get_json(silent=True) or {}
        entries = payload.get("counts") or []

        items = {item["id"]: item for item in build_inventory()}
        rows = []
        now = now_text()
        period = datetime.now().strftime("%m/%Y")

        for entry in entries:
            item = items.get(clean_value(entry.get("id")))
            if not item:
                continue

            counted = parse_quantity(entry.get("counted"), allow_blank=True)
            if counted is None:
                continue

            damaged = parse_quantity(entry.get("damaged"), allow_blank=True) or 0
            if damaged > counted:
                return json_error(f"{item['name']}: số hỏng ({damaged}) lớn hơn số đếm ({counted}).")

            rows.append({
                "Thời gian": now,
                "Kỳ": period,
                "Mã vật tư": item["id"],
                "Tên vật tư": item["name"],
                "SL đếm": counted,
                "SL hỏng": damaged,
                "Đang cho mượn lúc kiểm": item["on_loan"],
                "Người kiểm": session.get("user", ""),
                "Ghi chú": clean_value(entry.get("note"))
            })

        if not rows:
            return json_error("Chưa nhập số đếm cho món nào.")

        sheet, _, header = read_table(COUNT_SHEET, COUNT_HEADERS, fresh=True)
        append_records(sheet, COUNT_SHEET, COUNT_HEADERS, header, rows)

        return jsonify({"success": True, "saved": len(rows)})

    except ValueError as e:
        return json_error(str(e))
    except Exception as e:
        print("API INVENTORY COUNT ERROR:", e)
        return json_error(str(e), 500)


@app.route("/api/inventory/history")
def api_inventory_history():
    denied = require_admin_api()
    if denied:
        return denied

    try:
        item_id = clean_value(request.args.get("id"))
        _, counts, _ = read_table(COUNT_SHEET, COUNT_HEADERS)
        data = [
            {
                "at": count["Thời gian"],
                "counted": to_int(count["SL đếm"]),
                "damaged": to_int(count["SL hỏng"]),
                "on_loan": to_int(count["Đang cho mượn lúc kiểm"]),
                "by": count["Người kiểm"],
                "note": count["Ghi chú"]
            }
            for count in counts if count["Mã vật tư"] == item_id
        ]
        data.sort(key=lambda row: parse_datetime_any(row["at"]) or datetime.min, reverse=True)
        return jsonify({"success": True, "data": data})

    except Exception as e:
        print("API INVENTORY HISTORY ERROR:", e)
        return json_error(str(e), 500)


# ---------- Mượn – trả (món BÁO KTV) ----------

def shift_end(date_obj, shift):
    times = STAFF_SHIFT_TIMES.get(clean_value(shift).split(",")[-1].strip())
    if not date_obj or not times:
        return None
    hour, minute = times.split("-")[-1].strip().split(":")
    return date_obj.replace(hour=int(hour), minute=int(minute), second=0)


def handed_over_ids():
    return {
        clean_value(item.get("id_key"))
        for item in build_cleaning_records()
        if item.get("status") and item.get("status") != "Chưa bàn giao"
    }


def build_loans():
    _, loans, _ = read_table(LOAN_SHEET, LOAN_HEADERS)
    open_loans = [loan for loan in loans if loan["Trạng thái"] == LOAN_OPEN]
    handed = handed_over_ids() if open_loans and "ban_giao" in LOAN_RECALL_RULES else set()
    now = datetime.now()
    data = []

    for loan in loans:
        use_date = parse_datetime_any(loan["Ngày sử dụng"]) or parse_datetime_any(loan["Thời gian mượn"])
        reasons = []

        if loan["Trạng thái"] == LOAN_OPEN:
            if "qua_ngay" in LOAN_RECALL_RULES and use_date and use_date.date() < now.date():
                reasons.append(LOAN_RECALL_TEXT["qua_ngay"])
            if "ban_giao" in LOAN_RECALL_RULES and loan["ID_KEY"] in handed:
                reasons.append(LOAN_RECALL_TEXT["ban_giao"])
            end = shift_end(use_date, loan["Ca"])
            if "het_ca" in LOAN_RECALL_RULES and end and now > end:
                reasons.append(LOAN_RECALL_TEXT["het_ca"])

        data.append({
            "id": loan["Mã mượn"],
            "borrowed_at": loan["Thời gian mượn"],
            "booking_id": loan["ID_KEY"],
            "lab": loan["Phòng Lab"],
            "date": loan["Ngày sử dụng"],
            "date_iso": format_date_for_calendar(loan["Ngày sử dụng"]),
            "shift": loan["Ca"],
            "student_id": loan["MSSV"],
            "name": loan["Họ tên"],
            "item_id": loan["Mã vật tư"],
            "item_name": loan["Tên vật tư"],
            "qty": to_int(loan["Số lượng"]) or 0,
            "recorded_by": loan["Người ghi mượn"],
            "status": loan["Trạng thái"],
            "returned_at": loan["Thời gian trả"],
            "condition": loan["Tình trạng khi trả"],
            "received_by": loan["Người nhận trả"],
            "note": loan["Ghi chú"],
            "incident_id": loan["Mã sự cố"],
            "recall": bool(reasons),
            "recall_reasons": reasons
        })

    data.sort(key=lambda item: parse_datetime_any(item["borrowed_at"]) or datetime.min, reverse=True)
    return data


def loan_recall_notifications():
    _, loans, _ = read_table(LOAN_SHEET, LOAN_HEADERS)
    if not any(loan["Trạng thái"] == LOAN_OPEN for loan in loans):
        return []

    return [
        {
            "notification_id": f"loan-recall-{loan['id']}",
            "type": "loan_recall",
            "title": "CẦN THU HỒI VẬT TƯ",
            "message": f"{loan['name']} chưa trả {loan['qty']} {loan['item_name']} ({', '.join(loan['recall_reasons'])})",
            "id_key": loan["booking_id"],
            "date": loan["date"] or loan["borrowed_at"],
            "lab": loan["lab"],
            "level": "danger"
        }
        for loan in build_loans() if loan["recall"]
    ]


@app.route("/api/loans")
def api_loans():
    denied = require_admin_api()
    if denied:
        return denied

    try:
        data = [loan for loan in build_loans() if in_lab_scope(loan["lab"])]

        return jsonify({
            "success": True,
            "data": data,
            "recall_rules": [LOAN_RECALL_TEXT[rule] for rule in LOAN_RECALL_RULES if rule in LOAN_RECALL_TEXT],
            "conditions": RETURN_CONDITIONS
        })

    except Exception as e:
        print("API LOANS ERROR:", e)
        return json_error(str(e), 500)


@app.route("/api/loans/bookings")
def api_loan_bookings():
    """Lượt đăng ký của 1 ngày (mặc định hôm nay) + người trong lượt, để KTV chọn khi ghi mượn."""
    denied = require_admin_api()
    if denied:
        return denied

    try:
        date_iso = format_date_for_calendar(request.args.get("date")) or datetime.now().strftime("%Y-%m-%d")
        groups = {}

        for record in build_attendance_records():
            if format_date_for_calendar(record.get("date")) != date_iso:
                continue
            if not in_lab_scope(record.get("lab")):
                continue

            key = (record.get("booking_id"), record.get("shift"))

            if key not in groups:
                groups[key] = {
                    "booking_id": record.get("booking_id"),
                    "lab": normalize_lab_name(record.get("lab")),
                    "date": record.get("date"),
                    "shift": record.get("shift"),
                    "register_type": record.get("register_type") or "Cá nhân",
                    "people": []
                }

            groups[key]["people"].append({
                "student_id": record.get("student_id"),
                "name": record.get("name"),
                "role": record.get("role"),
                "attendance_status": record.get("attendance_status")
            })

        data = sorted(groups.values(), key=lambda item: (item["shift"], item["lab"], item["booking_id"]))
        return jsonify({"success": True, "date": date_iso, "data": data})

    except Exception as e:
        print("API LOAN BOOKINGS ERROR:", e)
        return json_error(str(e), 500)


@app.route("/api/loans", methods=["POST"])
def api_loan_add():
    denied = require_admin_api()
    if denied:
        return denied

    try:
        payload = request.get_json(silent=True) or {}
        booking_id = clean_value(payload.get("booking_id"))
        student_id = clean_value(payload.get("student_id"))
        name = clean_value(payload.get("name"))
        item_id = clean_value(payload.get("item_id"))
        qty = parse_quantity(payload.get("qty"))
        note = clean_value(payload.get("note"))

        if not booking_id:
            return json_error("Vui lòng chọn lượt đăng ký.")
        if not (student_id or name):
            return json_error("Vui lòng chọn người mượn.")
        if qty < 1:
            return json_error("Số lượng mượn phải từ 1 trở lên.")

        item = next((entry for entry in build_inventory() if entry["id"] == item_id), None)

        if not item or not item["active"]:
            return json_error("Vui lòng chọn vật tư trong danh mục.")
        if item["group"] != ITEM_GROUP_ASK:
            return json_error(f"{item['name']} thuộc nhóm TỰ LẤY – sinh viên tự lấy, không ghi mượn.")

        sheet, _, header = read_table(LOAN_SHEET, LOAN_HEADERS, fresh=True)
        loan_id = "MT-" + uuid.uuid4().hex[:8].upper()

        append_records(sheet, LOAN_SHEET, LOAN_HEADERS, header, [{
            "Mã mượn": loan_id,
            "Thời gian mượn": now_text(),
            "ID_KEY": booking_id,
            "Phòng Lab": normalize_lab_name(payload.get("lab")) or item["lab"],
            "Ngày sử dụng": clean_value(payload.get("date")),
            "Ca": clean_value(payload.get("shift")),
            "MSSV": student_id,
            "Họ tên": name,
            "Mã vật tư": item["id"],
            "Tên vật tư": item["name"],
            "Số lượng": qty,
            "Người ghi mượn": session.get("user", ""),
            "Trạng thái": LOAN_OPEN,
            "Ghi chú": note
        }])

        warning = ""
        if item["effective_qty"] is not None and item["on_loan"] + qty > item["effective_qty"]:
            warning = (f"Theo lần kiểm kê gần nhất chỉ có {item['effective_qty']} {item['unit']}, "
                       f"đang cho mượn {item['on_loan'] + qty}. Kiểm tra lại số trong kho.")

        return jsonify({"success": True, "id": loan_id, "warning": warning})

    except ValueError as e:
        return json_error(str(e))
    except Exception as e:
        print("API LOAN ADD ERROR:", e)
        return json_error(str(e), 500)


@app.route("/api/loans/return", methods=["POST"])
def api_loan_return():
    """KTV nhận trả: Bình thường / Hỏng / Thiếu. Hỏng hoặc Thiếu -> tự ghi Nhật ký sự cố."""
    denied = require_admin_api()
    if denied:
        return denied

    try:
        payload = request.get_json(silent=True) or {}
        loan_id = clean_value(payload.get("id"))
        condition = clean_value(payload.get("condition"))
        note = clean_value(payload.get("note"))

        if condition not in RETURN_CONDITIONS:
            return json_error("Vui lòng chọn tình trạng khi trả.")
        if condition != "Bình thường" and not note:
            return json_error("Vui lòng ghi rõ hỏng / thiếu như thế nào.")

        sheet, loans, header = read_table(LOAN_SHEET, LOAN_HEADERS, fresh=True)
        loan = next((row for row in loans if row["Mã mượn"] == loan_id), None)

        if not loan:
            return json_error("Không tìm thấy lượt mượn này.", 404)
        if loan["Trạng thái"] != LOAN_OPEN:
            return json_error("Lượt mượn này đã được nhận trả.")

        incident_id = ""

        if condition != "Bình thường":
            incident_sheet, _, incident_header = read_table(INCIDENT_SHEET, INCIDENT_HEADERS, fresh=True)
            incident_id = "SC-" + uuid.uuid4().hex[:8].upper()
            append_records(incident_sheet, INCIDENT_SHEET, INCIDENT_HEADERS, incident_header, [{
                "Mã sự cố": incident_id,
                "Thời gian ghi nhận": now_text(),
                "Nguồn": SOURCE_RETURN,
                "Mã nguồn": loan_id,
                "Phòng Lab": loan["Phòng Lab"],
                "Đối tượng": loan["Mã vật tư"],
                "Tên đối tượng": loan["Tên vật tư"],
                "Mô tả": f"{condition} khi trả ({loan['Số lượng']} món) – "
                         f"{loan['Họ tên']}{' · ' + loan['MSSV'] if loan['MSSV'] else ''}, lượt {loan['ID_KEY']}. {note}",
                "Người báo": session.get("user", ""),
                "Trạng thái": INCIDENT_NEW,
                "Cập nhật lúc": now_text()
            }])

        loan.update({
            "Trạng thái": LOAN_RETURNED,
            "Thời gian trả": now_text(),
            "Tình trạng khi trả": condition,
            "Người nhận trả": session.get("user", ""),
            "Ghi chú": " · ".join(part for part in [loan["Ghi chú"], note] if part),
            "Mã sự cố": incident_id
        })
        update_record(sheet, LOAN_SHEET, LOAN_HEADERS, header, loan)

        return jsonify({"success": True, "incident_id": incident_id})

    except Exception as e:
        print("API LOAN RETURN ERROR:", e)
        return json_error(str(e), 500)


# ---------- Nhật ký sự cố ----------

ABNORMAL_KEYS = [
    "10. Phòng Lab có phát sinh vấn đề bất thường trong quá trình sử dụng không?",
    "Phòng Lab có phát sinh vấn đề bất thường trong quá trình sử dụng không?",
    "15. Phòng Lab có gì bất thường sau khi sử dụng không?",
    "15. Phòng lab có gì bất thường sau khi sử dụng không?"
]


def handover_problem_description(row):
    text = get_value_by_possible_keys(row, ["GHI_CHU", "TINH_TRANG", "ERROR_MESSAGE"])

    if text:
        return text

    # Form cũ: câu mô tả vấn đề (nếu có) – tìm cột có chữ "mô tả" / "chi tiết"
    for key, value in row.items():
        key_norm = normalize_text(key)
        if ("mota" in key_norm or "chitiet" in key_norm) and clean_value(value):
            return clean_value(value)

    return "Phiếu bàn giao chọn \"Có sự cố\" (chưa có mô tả)."


def handover_incidents():
    """Sự cố tự lấy từ phiếu bàn giao chọn 'Có sự cố' (chưa lưu vào sheet)."""
    data = []

    for row in get_lab_handover_rows():
        if clean_value(row.get("BAN_GIAO_CHUNG_VOI")):
            continue   # cùng 1 phiếu áp cho nhiều ca -> chỉ lấy 1 lần

        id_key = get_value_by_possible_keys(row, ["1. ID_KEY của bạn là:", "ID_KEY", "ID_KEY của bạn là:"])
        abnormal = remove_vietnamese_accents(get_value_by_possible_keys(row, ABNORMAL_KEYS)).lower().strip()

        if not id_key or not abnormal.startswith("co"):
            continue

        lab = normalize_lab_name(
            get_value_by_possible_keys(row, ["7. Phòng Lab bạn đã sử dụng là:", "9. Phòng Lab bạn đã sử dụng là:", "Phòng", "Lab"])
            or extract_lab_from_id_key(id_key)
        )

        # Phiếu bàn giao mới (SOP-04): sinh viên chọn sẵn mã máy / vị trí và loại sự cố
        target = clean_value(row.get("SU_CO_VI_TRI")).upper()
        if target not in MACHINE_BY_CODE and target not in POSITION_BY_CODE:
            target = AREA_TARGET["code"]
        description = handover_problem_description(row)
        if clean_value(row.get("SU_CO_DA_DUNG_MAY")):
            description = f"{description} [SV xác nhận đã dừng máy, ngắt điện, báo KTV]".strip()

        data.append({
            "Mã sự cố": f"BG-{id_key}",
            "Thời gian ghi nhận": get_value_by_possible_keys(row, ["Dấu thời gian", "UPDATED_AT"]),
            "Nguồn": SOURCE_HANDOVER,
            "Mã nguồn": id_key,
            "Phòng Lab": lab,
            "Đối tượng": target,
            "Tên đối tượng": incident_target_name(target) if target != AREA_TARGET["code"] else AREA_TARGET["name"],
            "Mô tả": description,
            "Loại sự cố": clean_value(row.get("SU_CO_LOAI")) if clean_value(row.get("SU_CO_LOAI")) in INCIDENT_TYPES else "",
            "Người báo": get_value_by_possible_keys(row, ["3. Họ và tên của bạn là:", "5. Họ và tên của bạn là:", "Họ và tên"]),
            "Trạng thái": INCIDENT_NEW,
            "Cách xử lý": "",
            "Người xử lý": "",
            "Ngày xong": "",
            "Cập nhật lúc": "",
            "_row": None
        })

    return data


def build_incidents():
    _, stored, _ = read_table(INCIDENT_SHEET, INCIDENT_HEADERS)
    stored_sources = {(row["Nguồn"], row["Mã nguồn"]) for row in stored}

    try:
        virtual = [
            row for row in handover_incidents()
            if (row["Nguồn"], row["Mã nguồn"]) not in stored_sources
        ]
    except Exception as e:
        print("HANDOVER INCIDENTS ERROR:", e)
        virtual = []

    data = []

    for row in stored + virtual:
        data.append({
            "id": row["Mã sự cố"],
            "reported_at": row["Thời gian ghi nhận"],
            "source": row["Nguồn"],
            "source_id": row["Mã nguồn"],
            "lab": row["Phòng Lab"] or DEFAULT_LAB,
            "target": row["Đối tượng"],
            "target_name": row["Tên đối tượng"],
            "description": row["Mô tả"],
            "reporter": row["Người báo"],
            "status": row["Trạng thái"] or INCIDENT_NEW,
            "resolution": row["Cách xử lý"],
            "handled_by": row["Người xử lý"],
            "done_date": row["Ngày xong"],
            "updated_at": row["Cập nhật lúc"],
            "type": row.get("Loại sự cố", ""),
            "stored": row.get("_row") is not None
        })

    data.sort(key=lambda item: parse_datetime_any(item["reported_at"]) or datetime.min, reverse=True)
    return data


def machine_board(incidents):
    board = []

    for machine in MACHINES:
        open_items = [
            item for item in incidents
            if item["target"] == machine["code"] and item["status"] != INCIDENT_DONE
        ]
        if any(item["status"] == INCIDENT_NEW for item in open_items):
            state = "su_co"
        elif open_items:
            state = "dang_xu_ly"
        else:
            state = "binh_thuong"

        board.append(dict(machine, state=state, open_count=len(open_items)))

    return board


@app.route("/api/incidents")
def api_incidents():
    denied = require_admin_api()
    if denied:
        return denied

    try:
        incidents = [item for item in build_incidents() if in_lab_scope(item["lab"])]
        # Việc khắc phục 5S (SOP-05) liên kết với từng sự cố
        try:
            _, action_rows, _ = read_table(ACTION_SHEET, ACTION_HEADERS)
            linked = {}
            for row in action_rows:
                if row["Mã sự cố"]:
                    linked.setdefault(row["Mã sự cố"], []).append(row["Mã khắc phục"])
            for item in incidents:
                item["actions"] = linked.get(item["id"], [])
        except Exception as e:
            print("INCIDENT ACTION LINK ERROR:", e)
        items = [
            {"id": item["id"], "name": item["name"], "location": item["location"]}
            for item in build_inventory() if item["active"]
        ]

        return jsonify({
            "success": True,
            "data": incidents,
            "machines": machine_board(incidents),
            "items": items,
            "area": AREA_TARGET,
            "positions": POSITION_TARGETS,
            "types": INCIDENT_TYPES,
            "statuses": INCIDENT_STATUSES
        })

    except Exception as e:
        print("API INCIDENTS ERROR:", e)
        return json_error(str(e), 500)


def incident_target_name(code):
    if code in MACHINE_BY_CODE:
        return f"{code} · {MACHINE_BY_CODE[code]['name']}"
    if code in POSITION_BY_CODE:
        return POSITION_BY_CODE[code]["name"]
    if code == AREA_TARGET["code"]:
        return AREA_TARGET["name"]
    item = next((entry for entry in build_inventory() if entry["id"] == code), None)
    return item["name"] if item else ""


@app.route("/api/incidents", methods=["POST"])
def api_incident_add():
    """KTV tự thêm sự cố."""
    denied = require_admin_api()
    if denied:
        return denied

    try:
        payload = request.get_json(silent=True) or {}
        target = clean_value(payload.get("target")) or AREA_TARGET["code"]
        description = clean_value(payload.get("description"))
        target_name = incident_target_name(target)

        if not description:
            return json_error("Vui lòng mô tả sự cố.")
        if not target_name:
            return json_error("Đối tượng không hợp lệ.")

        sheet, _, header = read_table(INCIDENT_SHEET, INCIDENT_HEADERS, fresh=True)
        incident_id = "SC-" + uuid.uuid4().hex[:8].upper()

        append_records(sheet, INCIDENT_SHEET, INCIDENT_HEADERS, header, [{
            "Mã sự cố": incident_id,
            "Thời gian ghi nhận": now_text(),
            "Nguồn": SOURCE_STAFF,
            "Mã nguồn": "",
            "Phòng Lab": normalize_lab_name(payload.get("lab")) or DEFAULT_LAB,
            "Đối tượng": target,
            "Tên đối tượng": target_name,
            "Mô tả": description,
            "Người báo": clean_value(payload.get("reporter")) or session.get("user", ""),
            "Trạng thái": INCIDENT_NEW,
            "Cập nhật lúc": now_text(),
            "Loại sự cố": clean_value(payload.get("type")) if clean_value(payload.get("type")) in INCIDENT_TYPES else ""
        }])

        return jsonify({"success": True, "id": incident_id})

    except Exception as e:
        print("API INCIDENT ADD ERROR:", e)
        return json_error(str(e), 500)


@app.route("/api/incidents/update", methods=["POST"])
def api_incident_update():
    """Cập nhật cách xử lý / trạng thái / ngày xong. Sự cố từ phiếu bàn giao được lưu vào sheet ở lần cập nhật đầu."""
    denied = require_admin_api()
    if denied:
        return denied

    try:
        payload = request.get_json(silent=True) or {}
        incident_id = clean_value(payload.get("id"))
        status = clean_value(payload.get("status"))
        resolution = clean_value(payload.get("resolution"))
        done_date = clean_value(payload.get("done_date"))
        target = clean_value(payload.get("target"))

        if status not in INCIDENT_STATUSES:
            return json_error("Trạng thái không hợp lệ.")

        if status == INCIDENT_DONE:
            if not resolution:
                return json_error("Vui lòng ghi cách xử lý trước khi chuyển sang Đã xong.")
            done_date = format_date_for_calendar(done_date) or datetime.now().strftime("%Y-%m-%d")
            done_date = datetime.strptime(done_date, "%Y-%m-%d").strftime("%d/%m/%Y")
        else:
            done_date = ""

        sheet, rows, header = read_table(INCIDENT_SHEET, INCIDENT_HEADERS, fresh=True)
        record = next((row for row in rows if row["Mã sự cố"] == incident_id), None)

        if record is None:
            virtual = next((row for row in handover_incidents() if row["Mã sự cố"] == incident_id), None)
            if not virtual:
                return json_error("Không tìm thấy sự cố này.", 404)
            record = dict(virtual)

        record.update({
            "Trạng thái": status,
            "Cách xử lý": resolution,
            "Người xử lý": session.get("user", "") if (resolution or status != INCIDENT_NEW) else record.get("Người xử lý", ""),
            "Ngày xong": done_date,
            "Cập nhật lúc": now_text()
        })

        if target and target != record.get("Đối tượng"):
            name = incident_target_name(target)
            if name:
                record["Đối tượng"] = target
                record["Tên đối tượng"] = name

        incident_type = clean_value(payload.get("type"))
        if incident_type in INCIDENT_TYPES:
            record["Loại sự cố"] = incident_type

        if record.get("_row"):
            update_record(sheet, INCIDENT_SHEET, INCIDENT_HEADERS, header, record)
        else:
            append_records(sheet, INCIDENT_SHEET, INCIDENT_HEADERS, header, [record])

        return jsonify({"success": True})

    except Exception as e:
        print("API INCIDENT UPDATE ERROR:", e)
        return json_error(str(e), 500)


# =========================
# THEO DÕI RA VÀO: KTV BỔ SUNG SINH VIÊN VÀO LƯỢT ĐĂNG KÝ
# Nhóm đăng ký thiếu thành viên -> KTV thêm sinh viên vào đúng lượt (mã đăng ký + ngày + ca)
# để điểm danh. Dữ liệu đăng ký gốc không đổi; danh sách bổ sung lưu ở sheet riêng.
# =========================

EXTRA_MEMBER_SHEET = "LAB_ATTENDANCE_EXTRA"
EXTRA_MEMBER_HEADERS = [
    "Mã bổ sung", "Thời gian", "Mã đăng ký", "Lab", "Ngày", "Ca",
    "MSSV", "Họ tên", "Ghi chú", "Người thêm", "Trạng thái"
]
EXTRA_ACTIVE = "ĐANG HIỆU LỰC"
EXTRA_REMOVED = "ĐÃ XÓA"
EXTRA_ROLE = "Thành viên bổ sung (KTV)"


def extra_member_records(base_records, latest_status):
    try:
        _, rows, _ = read_table(EXTRA_MEMBER_SHEET, EXTRA_MEMBER_HEADERS)
    except Exception as e:
        print("EXTRA MEMBER READ ERROR:", e)
        return []

    data = []

    for row in rows:
        if row["Trạng thái"] == EXTRA_REMOVED:
            continue

        lead = next(
            (record for record in base_records
             if record["booking_id"] == row["Mã đăng ký"]
             and record["date"] == row["Ngày"]
             and record["shift"] == row["Ca"]),
            None
        )

        key = make_attendance_key(row["Mã đăng ký"], row["MSSV"], row["Họ tên"], row["Lab"], row["Ngày"], row["Ca"])

        data.append({
            "booking_id": row["Mã đăng ký"],
            "register_type": "Nhóm",
            "name": row["Họ tên"],
            "student_id": row["MSSV"],
            "lab": row["Lab"],
            "date": row["Ngày"],
            "shift": row["Ca"],
            "role": EXTRA_ROLE,
            "raw_members": lead["raw_members"] if lead else "",
            "attendance_status": latest_status.get(key, "Chưa điểm danh"),
            "extra_id": row["Mã bổ sung"],
            "extra_note": row["Ghi chú"],
            "extra_by": row["Người thêm"]
        })

    return data


@app.route("/api/attendance/extra-member", methods=["POST"])
def api_attendance_extra_add():
    denied = require_admin_api()
    if denied:
        return denied

    try:
        payload = request.get_json(silent=True) or {}
        booking_id = clean_value(payload.get("booking_id"))
        date = clean_value(payload.get("date"))
        shift = clean_value(payload.get("shift"))
        student_id = normalize_student_id(payload.get("student_id"))
        name = " ".join(clean_value(payload.get("name")).split())
        note = clean_value(payload.get("note"))

        if not booking_id or not date or not shift:
            return json_error("Vui lòng chọn lượt đăng ký.")
        if not re.fullmatch(r"\d{13}", student_id):
            return json_error("MSSV phải gồm đúng 13 chữ số.")
        if not name:
            return json_error("Vui lòng nhập họ tên sinh viên.")

        records = build_attendance_records()
        same_turn = [
            record for record in records
            if record["booking_id"] == booking_id and record["date"] == date and record["shift"] == shift
        ]

        if not same_turn:
            return json_error("Không tìm thấy lượt đăng ký này.", 404)

        if any(normalize_student_id(record["student_id"]) == student_id for record in same_turn):
            return json_error(f"MSSV {student_id} đã có trong lượt {booking_id} ({shift}).")

        sheet, _, header = read_table(EXTRA_MEMBER_SHEET, EXTRA_MEMBER_HEADERS, fresh=True)
        extra_id = "BS-" + uuid.uuid4().hex[:8].upper()

        append_records(sheet, EXTRA_MEMBER_SHEET, EXTRA_MEMBER_HEADERS, header, [{
            "Mã bổ sung": extra_id,
            "Thời gian": now_text(),
            "Mã đăng ký": booking_id,
            "Lab": same_turn[0]["lab"],
            "Ngày": date,
            "Ca": shift,
            "MSSV": student_id,
            "Họ tên": name,
            "Ghi chú": note,
            "Người thêm": session.get("user", ""),
            "Trạng thái": EXTRA_ACTIVE
        }])

        return jsonify({"success": True, "id": extra_id})

    except Exception as e:
        print("API EXTRA MEMBER ADD ERROR:", e)
        return json_error(str(e), 500)


@app.route("/api/attendance/extra-member", methods=["DELETE"])
def api_attendance_extra_remove():
    """Gỡ sinh viên KTV đã bổ sung nhầm (giữ dòng trong sheet, đánh dấu ĐÃ XÓA)."""
    denied = require_admin_api()
    if denied:
        return denied

    try:
        payload = request.get_json(silent=True) or {}
        extra_id = clean_value(payload.get("id"))

        sheet, rows, header = read_table(EXTRA_MEMBER_SHEET, EXTRA_MEMBER_HEADERS, fresh=True)
        row = next((item for item in rows if item["Mã bổ sung"] == extra_id), None)

        if not row or row["Trạng thái"] == EXTRA_REMOVED:
            return json_error("Không tìm thấy sinh viên bổ sung này.", 404)

        row["Trạng thái"] = EXTRA_REMOVED
        row["Ghi chú"] = " · ".join(part for part in [row["Ghi chú"], f"Gỡ bởi {session.get('user', '')} lúc {now_text()}"] if part)
        update_record(sheet, EXTRA_MEMBER_SHEET, EXTRA_MEMBER_HEADERS, header, row)

        return jsonify({"success": True})

    except Exception as e:
        print("API EXTRA MEMBER REMOVE ERROR:", e)
        return json_error(str(e), 500)


# =========================
# ĐÁNH GIÁ 5S ĐỊNH KỲ (SOP-05) – LIÊN KẾT VỚI BÀN GIAO, SỰ CỐ, VẬT TƯ, MƯỢN – TRẢ
# - KTV chấm 30 ô (6 khu A – F × S1 – S5, thang 0 – 4 hoặc KAD) ngay trên hệ thống, không dùng phiếu giấy.
# - Bên cạnh mỗi ô hiện bằng chứng hệ thống trong kỳ (bàn giao cuối ca, sự cố, kiểm kê, mượn – trả,
#   việc khắc phục đang mở, điểm lần trước). Bằng chứng chỉ để đối chiếu – điểm do người chấm quyết định.
# - Chốt phiếu: ô <= 2 tự tạo việc khắc phục (KHAC_PHUC); GV phụ trách xác nhận phiếu trên hệ thống.
# - Lỗi lặp lại khi bàn giao (SOP-02) cũng ghi vào KHAC_PHUC, cùng một danh sách theo dõi.
# =========================

FIVE_S = [
    {"s": "S1", "name": "Sàng lọc"},
    {"s": "S2", "name": "Sắp xếp"},
    {"s": "S3", "name": "Sạch sẽ"},
    {"s": "S4", "name": "Săn sóc"},
    {"s": "S5", "name": "Sẵn sàng"},
]
FIVE_S_CODES = [item["s"] for item in FIVE_S]

# 30 tiêu chí Lab 1 (6 khu × 5S) – soạn từ bảng kiểm cuối ca SOP-02 và yêu cầu chung SOP-05 mục 3.
# Bản đề xuất: KTV và GV phụ trách thống nhất trước lần chấm đầu tiên; sửa nội dung tại đây (một nguồn duy nhất).
AUDIT_ZONES = [
    {"zone": "A", "name": "Giảng viên & quản lý", "criteria": {
        "S1": "Bàn giảng viên không có đồ, dụng cụ của sinh viên; giấy tờ cũ đã dọn",
        "S2": "Máy chiếu, điều khiển, hồ sơ quản lý để đúng chỗ quy định",
        "S3": "Bàn, ghế, sàn khu A sạch, không bụi, rác",
        "S4": "Sơ đồ 5S, nội quy, SOP và tem QR niêm yết đúng chỗ, dễ đọc",
        "S5": "Lịch trực KTV cập nhật; phiếu bàn giao được KTV xác nhận trong ngày",
    }},
    {"zone": "B", "name": "Bàn B1 - B10", "criteria": {
        "S1": "Không còn giấy, chai nước, đồ cá nhân trên bàn và trong ngăn bàn",
        "S2": "Bàn đúng vạch chữ L vàng; 2 ghế / bàn, đẩy sát; không lấn lối đi",
        "S3": "Mặt bàn, ghế, sàn giữa các cụm sạch, không phoi, bụi",
        "S4": "Nhãn mã bàn và bảng kiểm cuối ca tại 5 cụm còn nguyên, dễ đọc",
        "S5": "Các lượt dùng bàn đều gửi phiếu bàn giao và được KTV xác nhận Đạt",
    }},
    {"zone": "C", "name": "Máy C1 - C7", "criteria": {
        "S1": "Không để phôi, sản phẩm, vật không dùng trên máy và quanh máy",
        "S2": "Phụ kiện, dao cụ cất đúng chỗ; lối tiếp cận, vùng an toàn quanh máy thông thoáng",
        "S3": "Máy và sàn quanh máy sạch, không phoi, bụi, dầu",
        "S4": "Decal mã C1 - C7 và tờ hướng dẫn vận hành tại máy còn nguyên, đúng máy",
        "S5": "Máy hỏng có biển ĐANG HỎNG; sự cố được ghi và xử lý đúng hạn SOP-04",
    }},
    {"zone": "D", "name": "Tủ T1, T2, kệ K1 - K3, D4", "criteria": {
        "S1": "Vật tư hỏng, không dùng đã tách khỏi tủ, kệ; tấm, thanh dài chỉ để ở D4",
        "S2": "Hộp, khay, rổ đúng mã vị trí; cửa tủ đóng, không để đồ trên nóc tủ",
        "S3": "Tủ, kệ, hộp sạch, không bụi, mạt",
        "S4": "Nhãn vị trí, nhãn xanh / đỏ, mục lục trong cánh tủ khớp danh mục trên hệ thống",
        "S5": "Kiểm kê trong tháng; hộp dưới tối thiểu đã bổ sung; món BÁO KTV được ghi mượn – trả",
    }},
    {"zone": "E", "name": "Khu ghế xếp", "criteria": {
        "S1": "Không để đồ khác chiếm khu ghế",
        "S2": "Ghế xếp gấp gọn, đúng khu E, không chắn cửa, lối đi",
        "S3": "Nền khu E sạch, không rác",
        "S4": "Vạch, biển khu E còn rõ",
        "S5": "Ghế hỏng được báo KTV và để riêng",
    }},
    {"zone": "F", "name": "Vệ sinh & chung cả phòng", "criteria": {
        "S1": "Không đồ cá nhân bỏ quên; thùng rác không đầy tràn",
        "S2": "Chổi, đồ hốt rác tại F1, máy hút bụi tại F2; lối thoát, bình chữa cháy không bị chắn",
        "S3": "Sàn toàn phòng sạch; thùng rác được đổ",
        "S4": "Sơ đồ 5S, SOP-01 - SOP-04, biển ĐANG HỎNG dùng đúng mẫu",
        "S5": "Đèn, quạt, điều hòa tắt cuối ngày; việc khắc phục 5S xử lý đúng hạn",
    }},
]
AUDIT_ZONE_BY_CODE = {item["zone"]: item for item in AUDIT_ZONES}
AUDIT_CELL_CODES = [f"{zone['zone']}-{s}" for zone in AUDIT_ZONES for s in FIVE_S_CODES]

# Mục có trong bảng kiểm cuối ca của cổng sinh viên (CHECKLIST_5S – SOP-02 trang 2), theo khu và S.
# Đồng bộ với portal: khu F luôn được kiểm, các khu khác chỉ khi sinh viên chọn đã dùng.
SHIFT_CHECK_CODES = {
    "A": {"S1", "S4"},
    "B": {"S1", "S2", "S3", "S4"},
    "C": {"S1", "S2", "S3", "S4", "S5"},
    "D": {"S1", "S2", "S5"},
    "E": {"S1", "S2", "S3"},
    "F": {"S1", "S2", "S3", "S5"},
}

# Mã máy / vị trí (sự cố) thuộc khu nào trên sơ đồ 5S Lab 1
AUDIT_ZONE_TARGETS = {
    "A": {"A"},
    "B": {f"B{i}" for i in range(1, 11)},
    "C": set(MACHINE_BY_CODE),
    "D": {"T1", "T2", "K1", "K2", "K3", "D4"},
    "E": {"E"},
    "F": {"F1", "F2", AREA_TARGET["code"]},
}

# Hệ thống chỉ lưu phiếu KTV tự đánh giá định kỳ (2 tuần / 1 lần). Đánh giá T0 / T1 / T2
# do nhóm nghiên cứu tự chấm bằng phiếu tay trong thời gian chạy thử, không nhập vào đây.
AUDIT_STAGE = "Định kỳ"
AUDIT_SCORES = ["0", "1", "2", "3", "4"]
AUDIT_KAD = "KAD"

AUDIT_DRAFT = "Nháp"
AUDIT_FINAL = "Đã chốt"
AUDIT_CONFIRMED = "GV đã xác nhận"
AUDIT_CANCELLED = "Đã hủy"
AUDIT_DONE_STATES = (AUDIT_FINAL, AUDIT_CONFIRMED)

AUDIT_DUE_DAYS = 14         # SOP-05: KTV chấm 2 tuần / 1 lần -> nhắc khi quá 14 ngày chưa có phiếu đã chốt
AUDIT_DEFAULT_PERIOD = 30   # kỳ dữ liệu mặc định khi chưa có phiếu trước

AUDIT_SHEET = "DANH_GIA_5S_PHIEU"
AUDIT_HEADERS = [
    "Mã đánh giá", "Phòng Lab", "Giai đoạn", "Ngày chấm", "Giờ bắt đầu", "Giờ kết thúc",
    "Người chấm", "Kỳ dữ liệu từ", "Kỳ dữ liệu đến", "Tổng điểm", "Số ô chấm", "Số ô KAD",
    "Điểm 5S (%)", "Xếp loại", "Trạng thái", "Số việc khắc phục", "GV xác nhận", "Xác nhận lúc",
    "Ghi chú", "Người tạo", "Cập nhật lúc"
]

AUDIT_CELL_SHEET = "DANH_GIA_5S"
AUDIT_CELL_HEADERS = [
    "Mã đánh giá", "Ngày chấm", "Phòng Lab", "Giai đoạn", "Người chấm", "Khu", "S", "Mã ô",
    "Tiêu chí", "Điểm", "Lý do KAD", "Vấn đề / ghi chú", "Minh chứng", "Bằng chứng hệ thống",
    "Cập nhật lúc"
]

ACTION_SHEET = "KHAC_PHUC"
ACTION_HEADERS = [
    "Mã khắc phục", "Ngày tạo", "Nguồn", "Mã nguồn", "Phòng Lab", "Khu", "S", "Mã lỗi",
    "Vấn đề", "Minh chứng", "Việc cần làm", "Người phụ trách", "Hạn", "Trạng thái", "Mã sự cố",
    "Kết quả", "Người xác nhận", "Ngày xác nhận", "Người tạo", "Cập nhật lúc"
]
ACTION_OPEN = "Chưa xử lý"
ACTION_PROGRESS = "Đang xử lý"
ACTION_DONE = "Đã khắc phục"
ACTION_STATUSES = [ACTION_OPEN, ACTION_PROGRESS, ACTION_DONE]
ACTION_SOURCE_AUDIT = "Đánh giá 5S định kỳ"
ACTION_SOURCE_REPEAT = "Lỗi lặp lại khi bàn giao"
ACTION_SOURCE_STAFF = "KTV ghi nhận"
ACTION_SOURCES = [ACTION_SOURCE_AUDIT, ACTION_SOURCE_REPEAT, ACTION_SOURCE_STAFF]
ACTION_DEFAULT_DAYS = 7     # xếp loại Đạt: khắc phục các ô <= 2 trong tuần

REPEAT_WINDOW_DAYS = 14     # lỗi lặp lại: cùng mã lỗi >= REPEAT_MIN lượt bàn giao trong 14 ngày
REPEAT_MIN = 3

# Hạn xử lý sự cố đề xuất (SOP-04 mục 8) – số ngày tính từ lúc ghi nhận
INCIDENT_DEADLINE_DAYS = {"Khẩn cấp": 0, "Hỏng thiết bị": 0, "Bất thường nhỏ": 2}

CELL_CODE_PATTERN = re.compile(r"^([A-F])-(S[1-5])$")


def audit_rating(percent):
    if percent is None:
        return ""
    if percent >= 85:
        return "Tốt"
    if percent >= 70:
        return "Đạt"
    return "Cần cải thiện"


def split_codes(text):
    return [
        code for code in dict.fromkeys(
            re.sub(r"\s+", "", part).upper()
            for part in clean_value(text).replace(";", ",").split(",")
        ) if CELL_CODE_PATTERN.match(code)
    ]


def iso_date(value):
    """Ngày từ mọi định dạng đang dùng (dd/mm/yyyy, yyyy-mm-dd, có giờ) -> date | None."""
    iso = format_date_for_calendar(value)
    try:
        return datetime.strptime(iso, "%Y-%m-%d").date() if iso else None
    except Exception:
        dt = parse_datetime_any(value)
        return dt.date() if dt else None


def vn_date(value):
    d = iso_date(value) if not hasattr(value, "strftime") else value
    return d.strftime("%d/%m/%Y") if d else ""


def audit_score_summary(cells):
    """cells: [{code, score}] -> tổng, số ô, KAD, %, xếp loại, % theo S và theo khu."""
    def pct(items):
        numbers = [int(item["score"]) for item in items if item.get("score") in AUDIT_SCORES]
        return round(sum(numbers) / (4 * len(numbers)) * 100, 1) if numbers else None

    numeric = [item for item in cells if item.get("score") in AUDIT_SCORES]
    total = sum(int(item["score"]) for item in numeric)
    percent = pct(cells)

    return {
        "total": total,
        "scored": len(numeric),
        "kad": len([item for item in cells if item.get("score") == AUDIT_KAD]),
        "blank": len([item for item in cells if not item.get("score")]),
        "percent": percent,
        "rating": audit_rating(percent),
        "by_s": {s: pct([c for c in cells if c["code"].endswith(s)]) for s in FIVE_S_CODES},
        "by_zone": {z["zone"]: pct([c for c in cells if c["code"].startswith(z["zone"] + "-")]) for z in AUDIT_ZONES},
        "low": [item["code"] for item in numeric if int(item["score"]) <= 2],
    }


# ---------- Đọc dữ liệu ----------

def read_audits(fresh=False):
    sheet, rows, header = read_table(AUDIT_SHEET, AUDIT_HEADERS, fresh=fresh)
    return sheet, rows, header


def read_audit_cells(fresh=False):
    return read_table(AUDIT_CELL_SHEET, AUDIT_CELL_HEADERS, fresh=fresh)


def audit_public(row, cells=None):
    """Dòng phiếu -> dict cho giao diện (tính lại điểm từ các ô nếu có)."""
    data = {
        "id": row["Mã đánh giá"],
        "lab": row["Phòng Lab"] or DEFAULT_LAB,
        "stage": row["Giai đoạn"],
        "date": row["Ngày chấm"],
        "date_iso": format_date_for_calendar(row["Ngày chấm"]),
        "start": row["Giờ bắt đầu"],
        "end": row["Giờ kết thúc"],
        "auditor": row["Người chấm"],
        "period_from": format_date_for_calendar(row["Kỳ dữ liệu từ"]),
        "period_to": format_date_for_calendar(row["Kỳ dữ liệu đến"]),
        "status": row["Trạng thái"] or AUDIT_DRAFT,
        "action_count": to_int(row["Số việc khắc phục"]) or 0,
        "confirmed_by": row["GV xác nhận"],
        "confirmed_at": row["Xác nhận lúc"],
        "note": row["Ghi chú"],
        "created_by": row["Người tạo"],
        "updated_at": row["Cập nhật lúc"],
    }

    if cells is not None:
        summary = audit_score_summary(cells)
    else:
        percent = None
        try:
            percent = float(row["Điểm 5S (%)"]) if row["Điểm 5S (%)"] != "" else None
        except ValueError:
            percent = None
        summary = {
            "total": to_int(row["Tổng điểm"]) or 0,
            "scored": to_int(row["Số ô chấm"]) or 0,
            "kad": to_int(row["Số ô KAD"]) or 0,
            "percent": percent,
            "rating": row["Xếp loại"] or audit_rating(percent),
        }

    data.update(summary)
    return data


def cells_of(audit_id, cell_rows):
    """Ô của một phiếu, đủ 30 ô theo cấu hình (ô chưa lưu -> trống)."""
    stored = {row["Mã ô"]: row for row in cell_rows if row["Mã đánh giá"] == audit_id}
    cells = []

    for zone in AUDIT_ZONES:
        for s in FIVE_S_CODES:
            code = f"{zone['zone']}-{s}"
            row = stored.get(code, {})
            cells.append({
                "code": code,
                "zone": zone["zone"],
                "s": s,
                # Phiếu đã lưu giữ đúng câu tiêu chí lúc chấm (sửa tiêu chí sau này không làm đổi phiếu cũ)
                "criterion": row.get("Tiêu chí") or zone["criteria"][s],
                "score": row.get("Điểm", ""),
                "kad_reason": row.get("Lý do KAD", ""),
                "note": row.get("Vấn đề / ghi chú", ""),
                "evidence": row.get("Minh chứng", ""),
                "system_evidence": row.get("Bằng chứng hệ thống", ""),
            })

    return cells


def finished_audits(lab=None, before=None, exclude=None):
    """Phiếu đã chốt (mới nhất trước); lọc theo phòng / trước ngày."""
    _, rows, _ = read_audits()
    items = []
    for row in rows:
        if row["Trạng thái"] not in AUDIT_DONE_STATES or row["Mã đánh giá"] == exclude:
            continue
        if lab and normalize_lab_name(row["Phòng Lab"] or DEFAULT_LAB) != normalize_lab_name(lab):
            continue
        d = iso_date(row["Ngày chấm"])
        if before and d and d > before:
            continue
        items.append((d or datetime.min.date(), row))
    items.sort(key=lambda item: (item[0], item[1]["Cập nhật lúc"]), reverse=True)
    return [row for _, row in items]


# ---------- Việc khắc phục ----------

def build_actions():
    _, rows, _ = read_table(ACTION_SHEET, ACTION_HEADERS)
    today = datetime.now().date()
    data = []

    try:
        incidents = {item["id"]: item for item in build_incidents()}
    except Exception as e:
        print("ACTION INCIDENT LINK ERROR:", e)
        incidents = {}

    for row in rows:
        due = iso_date(row["Hạn"])
        status = row["Trạng thái"] or ACTION_OPEN
        incident = incidents.get(row["Mã sự cố"]) if row["Mã sự cố"] else None
        data.append({
            "id": row["Mã khắc phục"],
            "created_at": row["Ngày tạo"],
            "source": row["Nguồn"],
            "source_id": row["Mã nguồn"],
            "lab": row["Phòng Lab"] or DEFAULT_LAB,
            "zone": row["Khu"],
            "s": row["S"],
            "code": row["Mã lỗi"],
            "problem": row["Vấn đề"],
            "evidence": row["Minh chứng"],
            "todo": row["Việc cần làm"],
            "owner": row["Người phụ trách"],
            "due": vn_date(due) if due else row["Hạn"],
            "due_iso": due.isoformat() if due else "",
            "status": status,
            "incident_id": row["Mã sự cố"],
            "incident_status": incident["status"] if incident else "",
            "incident_name": incident["target_name"] if incident else "",
            "result": row["Kết quả"],
            "confirmed_by": row["Người xác nhận"],
            "confirmed_at": row["Ngày xác nhận"],
            "created_by": row["Người tạo"],
            "updated_at": row["Cập nhật lúc"],
            "overdue": status != ACTION_DONE and bool(due) and due < today,
        })

    order = {ACTION_OPEN: 0, ACTION_PROGRESS: 1, ACTION_DONE: 2}
    data.sort(key=lambda item: (
        not item["overdue"], order.get(item["status"], 0), item["due_iso"] or "9999", item["id"]
    ))
    return data


def new_action_record(source, source_id, lab, code, problem, evidence, todo, owner, due, incident_id=""):
    zone, s = code.split("-") if CELL_CODE_PATTERN.match(code) else ("", "")
    return {
        "Mã khắc phục": "KP-" + uuid.uuid4().hex[:6].upper(),
        "Ngày tạo": now_text(),
        "Nguồn": source,
        "Mã nguồn": source_id,
        "Phòng Lab": lab or DEFAULT_LAB,
        "Khu": zone,
        "S": s,
        "Mã lỗi": code,
        "Vấn đề": problem,
        "Minh chứng": evidence,
        "Việc cần làm": todo,
        "Người phụ trách": owner,
        "Hạn": due,
        "Trạng thái": ACTION_OPEN,
        "Mã sự cố": incident_id,
        "Kết quả": "",
        "Người xác nhận": "",
        "Ngày xác nhận": "",
        "Người tạo": session.get("user", ""),
        "Cập nhật lúc": now_text(),
    }


# ---------- Bằng chứng hệ thống cho từng ô ----------

def handover_code_counts(lab, start, end):
    """Bàn giao cuối ca trong kỳ (theo ngày sử dụng): số lượt kiểm từng khu, lỗi tự khai / KTV ghi theo mã ô."""
    stats = {
        "booked": 0, "handed": 0, "ktv_done": 0, "ktv_pass": 0, "waiting_long": 0,
        "checked": {}, "self": {}, "ktv": {}, "any": {}
    }
    now = datetime.now()
    today = now.date()

    for record in build_cleaning_records():
        if normalize_lab_name(record.get("lab")) != normalize_lab_name(lab):
            continue
        d = iso_date(record.get("booking_date"))
        if not d or d < start or d > end:
            continue
        # Lượt hôm nay chưa bàn giao có thể chưa hết ca -> không tính là thiếu
        if record["status"] == "Chưa bàn giao" and d >= today:
            continue

        stats["booked"] += 1
        if record["status"] == "Chưa bàn giao":
            continue
        stats["handed"] += 1

        if not record.get("from_portal"):
            continue

        zones = {part.strip().upper() for part in (record.get("zones") or "").split(",") if part.strip()} | {"F"}
        for zone in zones:
            stats["checked"][zone] = stats["checked"].get(zone, 0) + 1

        self_codes = set(split_codes(record.get("error_codes")))
        ktv_codes = set(split_codes(record.get("ktv_codes")))
        for code in self_codes:
            stats["self"][code] = stats["self"].get(code, 0) + 1
        for code in ktv_codes:
            stats["ktv"][code] = stats["ktv"].get(code, 0) + 1
        for code in self_codes | ktv_codes:
            stats["any"][code] = stats["any"].get(code, 0) + 1

        if record.get("ktv_result"):
            stats["ktv_done"] += 1
            if record["ktv_result"] == KTV_PASS:
                stats["ktv_pass"] += 1

        if record.get("ktv_state") in (KTV_STATE_WAITING, KTV_STATE_RECHECK):
            handed_at = parse_datetime_any(record.get("redo_at") or record.get("handover_date"))
            if handed_at and now - handed_at > timedelta(hours=24):
                stats["waiting_long"] += 1

    return stats


def zone_of_target(code):
    code = clean_value(code).upper()
    if code.startswith("VT-"):
        return "D"
    for zone, targets in AUDIT_ZONE_TARGETS.items():
        if code in targets:
            return zone
    return ""


def audit_evidence(lab, start, end, exclude_audit=None):
    """Mã ô -> [{text, tone: ok|warn|bad, link}] từ dữ liệu hệ thống trong kỳ [start, end]."""
    evidence = {code: [] for code in AUDIT_CELL_CODES}
    lab = normalize_lab_name(lab) or DEFAULT_LAB
    today = datetime.now().date()

    def add(code, text, tone="ok", link=""):
        evidence[code].append({"text": text, "tone": tone, "link": link})

    def ratio_tone(bad, total, warn_at=0.0, bad_at=0.3):
        if not total or bad <= warn_at:
            return "ok"
        return "bad" if bad / total >= bad_at else "warn"

    # 1. Bàn giao cuối ca (SOP-01, SOP-02)
    try:
        stats = handover_code_counts(lab, start, end)
    except Exception as e:
        print("AUDIT EVIDENCE HANDOVER ERROR:", e)
        stats = None

    if stats:
        for zone, s_codes in SHIFT_CHECK_CODES.items():
            checked = stats["checked"].get(zone, 0)
            if not checked:
                continue
            for s in sorted(s_codes):
                code = f"{zone}-{s}"
                self_n = stats["self"].get(code, 0)
                ktv_n = stats["ktv"].get(code, 0)
                worst = max(self_n, ktv_n)
                add(code,
                    f"Bàn giao: {checked} lượt kiểm khu {zone}; SV tự khai lỗi {self_n}, KTV ghi lỗi {ktv_n}",
                    ratio_tone(worst, checked),
                    url_for("cleaning_report"))

        if stats["booked"]:
            missing = stats["booked"] - stats["handed"]
            add("B-S5", f"Đã bàn giao {stats['handed']}/{stats['booked']} lượt đăng ký",
                ratio_tone(missing, stats["booked"], bad_at=0.2), url_for("cleaning_report"))
        if stats["ktv_done"]:
            add("B-S5", f"KTV xác nhận Đạt {stats['ktv_pass']}/{stats['ktv_done']} phiếu bàn giao",
                ratio_tone(stats["ktv_done"] - stats["ktv_pass"], stats["ktv_done"]), url_for("cleaning_report"))
        if stats["waiting_long"]:
            add("A-S5", f"{stats['waiting_long']} phiếu bàn giao chờ KTV xác nhận quá 24 giờ", "warn",
                url_for("cleaning_report", ktv="cho_ktv"))

    # 2. Sự cố (SOP-04)
    try:
        incidents = [item for item in build_incidents() if normalize_lab_name(item["lab"]) == lab]
    except Exception as e:
        print("AUDIT EVIDENCE INCIDENT ERROR:", e)
        incidents = []

    by_zone = {}
    for item in incidents:
        zone = zone_of_target(item["target"])
        if not zone:
            continue
        reported = parse_datetime_any(item["reported_at"])
        in_period = reported and start <= reported.date() <= end
        is_open = item["status"] != INCIDENT_DONE
        if not (in_period or is_open):
            continue
        entry = by_zone.setdefault(zone, {"period": 0, "open": 0, "late": 0, "damage": 0})
        entry["period"] += 1 if in_period else 0
        if is_open:
            entry["open"] += 1
            limit = INCIDENT_DEADLINE_DAYS.get(item.get("type"), 2)
            if reported and (today - reported.date()).days > limit:
                entry["late"] += 1
            if item["target"].startswith("VT-"):
                entry["damage"] += 1

    for zone, entry in by_zone.items():
        tone = "bad" if entry["late"] else ("warn" if entry["open"] else "ok")
        text = f"Sự cố: {entry['period']} ghi nhận trong kỳ, {entry['open']} chưa xong"
        if entry["late"]:
            text += f", {entry['late']} quá hạn SOP-04"
        add(f"{zone}-S5", text, tone, url_for("inventory_page", tab="su-co"))
        if entry["damage"]:
            add("D-S1", f"{entry['damage']} món vật tư hỏng chưa xử lý xong", "warn",
                url_for("inventory_page", tab="su-co"))

    # 3. Vật tư, kiểm kê (SOP-03)
    try:
        items = [item for item in build_inventory()
                 if item["active"] and normalize_lab_name(item["lab"]) == lab]
    except Exception as e:
        print("AUDIT EVIDENCE INVENTORY ERROR:", e)
        items = []

    if items:
        total = len(items)
        not_counted = len([i for i in items if "chua_kiem_ke" in i["flags"] or "qua_han" in i["flags"]])
        below = len([i for i in items if "duoi_toi_thieu" in i["flags"]])
        reported_low = len([i for i in items if "sv_bao_het" in i["flags"]])
        damaged = len([i for i in items if "co_hong" in i["flags"]])
        bad_location = len([i for i in items if not LOCATION_PATTERN.match((i["location"] or "").upper())])

        add("D-S5", f"Kiểm kê trong {INVENTORY_COUNT_DAYS} ngày: {total - not_counted}/{total} món",
            ratio_tone(not_counted, total, bad_at=0.2), url_for("inventory_page"))
        if below or reported_low:
            add("D-S5", f"{below} món dưới tối thiểu, {reported_low} hộp sinh viên báo sắp hết chưa bổ sung",
                "bad" if below else "warn", url_for("inventory_page"))
        if damaged:
            add("D-S1", f"{damaged} món có số hỏng khi kiểm kê / sự cố chưa xong", "warn", url_for("inventory_page"))
        add("D-S4",
            f"{bad_location} món chưa có mã vị trí đúng quy tắc (T/K-N-ô)" if bad_location
            else f"Cả {total} món có mã vị trí đúng quy tắc",
            "warn" if bad_location else "ok", url_for("inventory_page"))
    else:
        add("D-S5", "Danh mục vật tư trên hệ thống còn trống – chưa đối chiếu được", "warn", url_for("inventory_page"))

    try:
        recall = [loan for loan in build_loans()
                  if loan["recall"] and normalize_lab_name(loan["lab"] or DEFAULT_LAB) == lab]
    except Exception as e:
        print("AUDIT EVIDENCE LOAN ERROR:", e)
        recall = []
    if recall:
        add("D-S5", f"{len(recall)} lượt món BÁO KTV cần thu hồi", "bad", url_for("loans_page"))

    # 4. Việc khắc phục đang mở
    try:
        actions = [a for a in build_actions() if normalize_lab_name(a["lab"]) == lab and a["status"] != ACTION_DONE]
    except Exception as e:
        print("AUDIT EVIDENCE ACTION ERROR:", e)
        actions = []

    for code in AUDIT_CELL_CODES:
        mine = [a for a in actions if a["code"] == code]
        if mine:
            late = len([a for a in mine if a["overdue"]])
            add(code, f"Khắc phục: {len(mine)} việc chưa xong" + (f" ({late} quá hạn)" if late else ""),
                "bad" if late else "warn", url_for("audit_page", tab="khac-phuc", q=code))

    late_all = len([a for a in actions if a["overdue"]])
    if late_all:
        add("F-S5", f"Toàn phòng: {late_all} việc khắc phục 5S quá hạn", "bad",
            url_for("audit_page", tab="khac-phuc"))

    # 5. Điểm lần trước (cùng phòng)
    try:
        previous = finished_audits(lab=lab, exclude=exclude_audit)
        if previous:
            prev = previous[0]
            _, cell_rows, _ = read_audit_cells()
            prev_cells = {c["code"]: c for c in cells_of(prev["Mã đánh giá"], cell_rows)}
            for code in AUDIT_CELL_CODES:
                score = prev_cells.get(code, {}).get("score")
                if score:
                    add(code, f"Lần trước ({prev['Ngày chấm']}): {score}",
                        "bad" if score in ("0", "1", "2") else "ok",
                        url_for("audit_page", id=prev["Mã đánh giá"]))
    except Exception as e:
        print("AUDIT EVIDENCE PREVIOUS ERROR:", e)

    return evidence


def evidence_text(items):
    return " | ".join(item["text"] for item in items)


def default_period(lab, audit_date, exclude=None):
    previous = [row for row in finished_audits(lab=lab, exclude=exclude)
                if iso_date(row["Ngày chấm"]) and iso_date(row["Ngày chấm"]) < audit_date]
    start = iso_date(previous[0]["Ngày chấm"]) if previous else audit_date - timedelta(days=AUDIT_DEFAULT_PERIOD)
    return start, audit_date


# ---------- Thông báo (chuông + Trang chủ) ----------

def audit_notifications():
    items = []
    today = datetime.now().date()

    try:
        done = finished_audits(lab=DEFAULT_LAB)
        last_date = iso_date(done[0]["Ngày chấm"]) if done else None
        if not last_date or (today - last_date).days >= AUDIT_DUE_DAYS:
            year, week, _ = today.isocalendar()
            items.append({
                "notification_id": f"audit-due-{year}-W{week}",
                "type": "audit_due",
                "title": "ĐẾN HẠN CHẤM 5S ĐỊNH KỲ",
                "message": (
                    f"Lab 1: lần chấm gần nhất {vn_date(last_date)} ({(today - last_date).days} ngày trước). "
                    "SOP-05: KTV chấm 2 tuần / 1 lần." if last_date
                    else "Lab 1 chưa có phiếu đánh giá 5S nào trên hệ thống – KTV chấm lần đầu theo SOP-05."
                ),
                "id_key": "",
                "date": now_text(),
                "lab": DEFAULT_LAB,
                "level": "warning",
                "link": url_for("audit_page", new=1),
                "link_label": "Chấm 5S"
            })
    except Exception as e:
        print("AUDIT DUE NOTIFICATION ERROR:", e)

    try:
        actions = build_actions()
    except Exception as e:
        print("ACTION NOTIFICATION ERROR:", e)
        actions = []

    for action in actions:
        if not action["overdue"]:
            continue
        items.append({
            "notification_id": f"action-late-{action['id']}-{action['due_iso']}",
            "type": "action_overdue",
            "title": "VIỆC KHẮC PHỤC 5S QUÁ HẠN",
            "message": f"{action['code'] or action['zone']} · {action['problem'] or action['todo']} – hạn {action['due']}"
                       + (f", phụ trách {action['owner']}" if action["owner"] else ""),
            "id_key": action["id"],
            "date": action["due"],
            "lab": action["lab"],
            "level": "danger",
            "link": url_for("audit_page", tab="khac-phuc", q=action["id"]),
            "link_label": "Xem việc"
        })

    # Lỗi lặp lại khi bàn giao (SOP-02 trang 2: lỗi lặp lại ghi KHAC_PHUC)
    try:
        start = today - timedelta(days=REPEAT_WINDOW_DAYS)
        stats = handover_code_counts(DEFAULT_LAB, start, today)
        open_codes = {a["code"] for a in actions if a["status"] != ACTION_DONE}
        for code, count in sorted(stats["any"].items(), key=lambda pair: -pair[1]):
            if count < REPEAT_MIN or code in open_codes:
                continue
            zone, s = code.split("-")
            items.append({
                "notification_id": f"repeat-{code}-{today.isoformat()}",
                "type": "repeat_error",
                "title": "LỖI 5S LẶP LẠI KHI BÀN GIAO",
                "message": f"{code} ({AUDIT_ZONE_BY_CODE[zone]['name']} – {s}): {count} lượt bàn giao trong {REPEAT_WINDOW_DAYS} ngày. Tạo việc khắc phục (SOP-02, SOP-05).",
                "id_key": code,
                "date": now_text(),
                "lab": DEFAULT_LAB,
                "level": "warning",
                "link": url_for("audit_page", tab="khac-phuc", code=code, n=count),
                "link_label": "Tạo việc"
            })
    except Exception as e:
        print("REPEAT ERROR NOTIFICATION ERROR:", e)

    return items


# ---------- Trang + API ----------

def audit_page_guard():
    if "user" not in session:
        return redirect(url_for("login"))
    return None


def require_login_api():
    if "user" not in session:
        return json_error("Phiên đăng nhập đã hết. Vui lòng đăng nhập lại.", 401)
    return None


@app.route("/danh-gia-5s")
def audit_page():
    # KTV: chấm, khắc phục. Giảng viên: xem và xác nhận phiếu (SOP-05 bước 8).
    return audit_page_guard() or render_template(
        "audit_5s.html", user=session["user"], is_admin=is_admin_session()
    )


@app.route("/danh-gia-5s/in/<audit_id>")
def audit_print(audit_id):
    denied = audit_page_guard()
    if denied:
        return denied

    _, rows, _ = read_audits()
    row = next((item for item in rows if item["Mã đánh giá"] == audit_id), None)
    if not row:
        return "Không tìm thấy phiếu đánh giá này.", 404

    _, cell_rows, _ = read_audit_cells()
    cells = cells_of(audit_id, cell_rows)
    audit = audit_public(row, cells)
    actions = [a for a in build_actions() if a["source_id"] == audit_id]

    return render_template(
        "audit_5s_print.html", audit=audit, cells={c["code"]: c for c in cells},
        zones=AUDIT_ZONES, five_s=FIVE_S, actions=actions,
        kad_cells=[c for c in cells if c["score"] == AUDIT_KAD]
    )


@app.route("/api/5s/config")
def api_audit_config():
    denied = require_login_api()
    if denied:
        return denied
    return jsonify({
        "success": True,
        "zones": AUDIT_ZONES,
        "five_s": FIVE_S,
        "action_statuses": ACTION_STATUSES,
        "action_sources": ACTION_SOURCES,
        "action_days": ACTION_DEFAULT_DAYS,
        "due_days": AUDIT_DUE_DAYS,
        "is_admin": is_admin_session(),
        "user": session.get("user", "")
    })


@app.route("/api/5s/audits")
def api_audits():
    denied = require_login_api()
    if denied:
        return denied

    try:
        _, rows, _ = read_audits()
        _, cell_rows, _ = read_audit_cells()
        data = [
            audit_public(row, cells_of(row["Mã đánh giá"], cell_rows))
            for row in rows
            if row["Trạng thái"] != AUDIT_CANCELLED and in_lab_scope(row["Phòng Lab"])
        ]
        data.sort(key=lambda item: (item["date_iso"] or "", item["updated_at"]), reverse=True)

        done = [item for item in data if item["status"] in AUDIT_DONE_STATES]
        last = done[0] if done else None
        days = (datetime.now().date() - iso_date(last["date"])).days if last and iso_date(last["date"]) else None

        return jsonify({
            "success": True,
            "data": data,
            "last": last,
            "days_since_last": days,
            "due": last is None or (days is not None and days >= AUDIT_DUE_DAYS)
        })

    except Exception as e:
        print("API AUDITS ERROR:", e)
        return json_error(str(e), 500)


@app.route("/api/5s/audits/<audit_id>")
def api_audit_detail(audit_id):
    denied = require_login_api()
    if denied:
        return denied

    try:
        _, rows, _ = read_audits()
        row = next((item for item in rows if item["Mã đánh giá"] == audit_id), None)
        if not row:
            return json_error("Không tìm thấy phiếu đánh giá này.", 404)

        _, cell_rows, _ = read_audit_cells()
        cells = cells_of(audit_id, cell_rows)
        audit = audit_public(row, cells)

        start = iso_date(row["Kỳ dữ liệu từ"])
        end = iso_date(row["Kỳ dữ liệu đến"])
        if not (start and end):
            start, end = default_period(audit["lab"], iso_date(row["Ngày chấm"]) or datetime.now().date(), audit_id)

        # Phiếu nháp: bằng chứng tính lại theo dữ liệu hiện tại; phiếu đã chốt: giữ bản đã lưu lúc chốt
        live = audit["status"] == AUDIT_DRAFT
        evidence = audit_evidence(audit["lab"], start, end, exclude_audit=audit_id) if live else {}

        actions = [a for a in build_actions() if a["source_id"] == audit_id]

        return jsonify({
            "success": True,
            "audit": audit,
            "cells": cells,
            "evidence": evidence,
            "evidence_live": live,
            "period": {"from": start.isoformat(), "to": end.isoformat()},
            "actions": actions
        })

    except Exception as e:
        print("API AUDIT DETAIL ERROR:", e)
        return json_error(str(e), 500)


@app.route("/api/5s/evidence")
def api_audit_evidence():
    denied = require_admin_api()
    if denied:
        return denied

    try:
        lab = normalize_lab_name(request.args.get("lab")) or DEFAULT_LAB
        audit_date = iso_date(request.args.get("date")) or datetime.now().date()
        start = iso_date(request.args.get("from"))
        end = iso_date(request.args.get("to"))
        if not (start and end):
            start, end = default_period(lab, audit_date, request.args.get("exclude"))
        if start > end:
            return json_error("Kỳ dữ liệu: ngày bắt đầu phải trước ngày kết thúc.")

        return jsonify({
            "success": True,
            "evidence": audit_evidence(lab, start, end, exclude_audit=request.args.get("exclude")),
            "period": {"from": start.isoformat(), "to": end.isoformat()}
        })

    except Exception as e:
        print("API AUDIT EVIDENCE ERROR:", e)
        return json_error(str(e), 500)


def write_audit_cells(audit_id, meta, cells, evidence):
    """Ghi 30 ô của một phiếu: cập nhật cả khối nếu đã có, chưa có thì thêm mới."""
    sheet, rows, header = read_audit_cells(fresh=True)
    existing = sorted([row for row in rows if row["Mã đánh giá"] == audit_id], key=lambda row: row["_row"])
    by_code = {row["Mã ô"]: row for row in existing}
    stamp = now_text()

    records = []
    for cell in cells:
        zone = AUDIT_ZONE_BY_CODE[cell["zone"]]
        records.append({
            "Mã đánh giá": audit_id,
            "Ngày chấm": meta["Ngày chấm"],
            "Phòng Lab": meta["Phòng Lab"],
            "Giai đoạn": meta["Giai đoạn"],
            "Người chấm": meta["Người chấm"],
            "Khu": cell["zone"],
            "S": cell["s"],
            "Mã ô": cell["code"],
            "Tiêu chí": by_code.get(cell["code"], {}).get("Tiêu chí") or zone["criteria"][cell["s"]],
            "Điểm": cell["score"],
            "Lý do KAD": cell["kad_reason"] if cell["score"] == AUDIT_KAD else "",
            "Vấn đề / ghi chú": cell["note"],
            "Minh chứng": cell["evidence"],
            "Bằng chứng hệ thống": evidence_text(evidence.get(cell["code"], [])) if evidence else
                                    by_code.get(cell["code"], {}).get("Bằng chứng hệ thống", ""),
            "Cập nhật lúc": stamp,
        })

    ensure_sheet_headers(sheet, AUDIT_CELL_HEADERS, header)
    contiguous = (
        len(existing) == len(records)
        and [row["Mã ô"] for row in existing] == [record["Mã ô"] for record in records]
        and existing[-1]["_row"] - existing[0]["_row"] == len(existing) - 1
    )

    if contiguous:
        sheet.update(
            range_name=f"A{existing[0]['_row']}",
            values=[[record.get(h, "") for h in AUDIT_CELL_HEADERS] for record in records],
            value_input_option="RAW"
        )
    else:
        missing = []
        for record in records:
            row = by_code.get(record["Mã ô"])
            if row:
                sheet.update(range_name=f"A{row['_row']}",
                             values=[[record.get(h, "") for h in AUDIT_CELL_HEADERS]],
                             value_input_option="RAW")
            else:
                missing.append(record)
        if missing:
            sheet.append_rows([[record.get(h, "") for h in AUDIT_CELL_HEADERS] for record in missing],
                              value_input_option="RAW")

    invalidate_cache(AUDIT_CELL_SHEET)


def parse_audit_cells(payload_cells):
    given = {clean_value(item.get("code")).upper(): item for item in (payload_cells or []) if isinstance(item, dict)}
    cells = []
    for zone in AUDIT_ZONES:
        for s in FIVE_S_CODES:
            code = f"{zone['zone']}-{s}"
            item = given.get(code, {})
            score = clean_value(item.get("score")).upper()
            if score not in AUDIT_SCORES + [AUDIT_KAD, ""]:
                raise ValueError(f"Ô {code}: điểm không hợp lệ ({score}).")
            cells.append({
                "code": code, "zone": zone["zone"], "s": s, "score": score,
                "kad_reason": clean_value(item.get("kad_reason"))[:300],
                "note": clean_value(item.get("note"))[:300],
                "evidence": clean_value(item.get("evidence"))[:300],
            })
    return cells


@app.route("/api/5s/audits", methods=["POST"])
def api_audit_save():
    """Lưu nháp / chốt phiếu đánh giá. Chốt: đủ 30 ô, KAD có lý do, ô <= 2 có vấn đề + minh chứng
    -> khóa phiếu, lưu bằng chứng hệ thống tại thời điểm chốt, tạo việc khắc phục cho ô <= 2."""
    denied = require_admin_api()
    if denied:
        return denied

    try:
        payload = request.get_json(silent=True) or {}
        audit_id = clean_value(payload.get("id"))
        finalize = bool(payload.get("finalize"))
        lab = normalize_lab_name(payload.get("lab")) or DEFAULT_LAB
        stage = AUDIT_STAGE
        audit_date = iso_date(payload.get("date"))
        start_time = clean_value(payload.get("start"))[:5]
        end_time = clean_value(payload.get("end"))[:5]
        auditor = clean_value(payload.get("auditor")) or session.get("user", "")
        period_from = iso_date(payload.get("period_from"))
        period_to = iso_date(payload.get("period_to"))

        if not audit_date:
            return json_error("Chọn ngày chấm.")
        if audit_date > datetime.now().date():
            return json_error("Ngày chấm không được sau hôm nay (SOP-05: ghi ngày thực tế, không điền lùi / trước).")
        if not (period_from and period_to):
            period_from, period_to = default_period(lab, audit_date, audit_id)
        if period_from > period_to:
            return json_error("Kỳ dữ liệu: ngày bắt đầu phải trước ngày kết thúc.")

        try:
            cells = parse_audit_cells(payload.get("cells"))
        except ValueError as error:
            return json_error(str(error))

        if finalize:
            errors = []
            blank = [c["code"] for c in cells if not c["score"]]
            if blank:
                errors.append(f"Còn {len(blank)} ô chưa chấm: {', '.join(blank[:8])}{'…' if len(blank) > 8 else ''}.")
            for c in cells:
                if c["score"] == AUDIT_KAD and not c["kad_reason"]:
                    errors.append(f"Ô {c['code']} chọn KAD: ghi lý do không áp dụng.")
                if c["score"] in ("0", "1", "2"):
                    if not c["note"]:
                        errors.append(f"Ô {c['code']} ({c['score']} điểm): ghi vấn đề cần khắc phục.")
                    if not c["evidence"]:
                        errors.append(f"Ô {c['code']} ({c['score']} điểm): ghi mã ảnh / link minh chứng.")
            if not start_time:
                errors.append("Ghi giờ bắt đầu chấm.")
            if errors:
                return jsonify({"success": False, "error": "Chưa chốt được phiếu", "errors": errors}), 400

        sheet, rows, header = read_audits(fresh=True)
        row = next((item for item in rows if item["Mã đánh giá"] == audit_id), None) if audit_id else None

        if row and row["Trạng thái"] != AUDIT_DRAFT:
            return json_error("Phiếu đã chốt – không sửa được. Tạo phiếu mới nếu cần chấm lại.")

        if not row:
            audit_id = f"DG5S-{audit_date.strftime('%Y%m%d')}-{uuid.uuid4().hex[:4].upper()}"

        summary = audit_score_summary(cells)
        evidence = audit_evidence(lab, period_from, period_to, exclude_audit=audit_id) if finalize else {}

        record = dict(row or {})
        record.update({
            "Mã đánh giá": audit_id,
            "Phòng Lab": lab,
            "Giai đoạn": stage,
            "Ngày chấm": audit_date.strftime("%d/%m/%Y"),
            "Giờ bắt đầu": start_time,
            "Giờ kết thúc": end_time or (datetime.now().strftime("%H:%M") if finalize else ""),
            "Người chấm": auditor,
            "Kỳ dữ liệu từ": period_from.strftime("%d/%m/%Y"),
            "Kỳ dữ liệu đến": period_to.strftime("%d/%m/%Y"),
            "Tổng điểm": summary["total"],
            "Số ô chấm": summary["scored"],
            "Số ô KAD": summary["kad"],
            "Điểm 5S (%)": summary["percent"] if summary["percent"] is not None else "",
            "Xếp loại": summary["rating"],
            "Trạng thái": AUDIT_FINAL if finalize else AUDIT_DRAFT,
            "Ghi chú": clean_value(payload.get("note"))[:500],
            "Người tạo": (row or {}).get("Người tạo") or session.get("user", ""),
            "Cập nhật lúc": now_text(),
        })

        write_audit_cells(audit_id, record, cells, evidence)

        created = []
        if finalize:
            a_sheet, a_rows, a_header = read_table(ACTION_SHEET, ACTION_HEADERS, fresh=True)
            have = {(r["Mã nguồn"], r["Mã lỗi"]) for r in a_rows if r["Nguồn"] == ACTION_SOURCE_AUDIT}
            due = (audit_date + timedelta(days=ACTION_DEFAULT_DAYS)).strftime("%d/%m/%Y")
            for c in cells:
                if c["score"] in ("0", "1", "2") and (audit_id, c["code"]) not in have:
                    zone = AUDIT_ZONE_BY_CODE[c["zone"]]
                    created.append(new_action_record(
                        ACTION_SOURCE_AUDIT, audit_id, lab, c["code"],
                        f"[{c['score']} điểm] {c['note']}", c["evidence"],
                        f"Khắc phục để đạt: {zone['criteria'][c['s']]}",
                        auditor, due
                    ))
            if created:
                append_records(a_sheet, ACTION_SHEET, ACTION_HEADERS, a_header, created)
            record["Số việc khắc phục"] = len(created) + len(
                [r for r in a_rows if r["Nguồn"] == ACTION_SOURCE_AUDIT and r["Mã nguồn"] == audit_id]
            )

        if row:
            update_record(sheet, AUDIT_SHEET, AUDIT_HEADERS, header, record)
        else:
            append_records(sheet, AUDIT_SHEET, AUDIT_HEADERS, header, [record])

        return jsonify({
            "success": True,
            "id": audit_id,
            "status": record["Trạng thái"],
            "percent": summary["percent"],
            "rating": summary["rating"],
            "actions_created": [r["Mã khắc phục"] for r in created]
        })

    except Exception as e:
        print("API AUDIT SAVE ERROR:", e)
        return json_error(str(e), 500)


@app.route("/api/5s/audits/confirm", methods=["POST"])
def api_audit_confirm():
    """GV phụ trách (hoặc KTV ghi thay tên GV) xác nhận phiếu đã chốt – thay chữ ký trên phiếu giấy."""
    denied = require_login_api()
    if denied:
        return denied

    try:
        payload = request.get_json(silent=True) or {}
        audit_id = clean_value(payload.get("id"))
        confirmer = clean_value(payload.get("confirmer")) or session.get("user", "")

        sheet, rows, header = read_audits(fresh=True)
        row = next((item for item in rows if item["Mã đánh giá"] == audit_id), None)
        if not row:
            return json_error("Không tìm thấy phiếu đánh giá này.", 404)
        if row["Trạng thái"] != AUDIT_FINAL:
            return json_error("Chỉ xác nhận được phiếu đã chốt và chưa xác nhận.")

        row["Trạng thái"] = AUDIT_CONFIRMED
        row["GV xác nhận"] = confirmer if is_admin_session() else f"{confirmer} (GV)"
        row["Xác nhận lúc"] = now_text()
        row["Cập nhật lúc"] = now_text()
        update_record(sheet, AUDIT_SHEET, AUDIT_HEADERS, header, row)

        return jsonify({"success": True})

    except Exception as e:
        print("API AUDIT CONFIRM ERROR:", e)
        return json_error(str(e), 500)


@app.route("/api/5s/audits/cancel", methods=["POST"])
def api_audit_cancel():
    """Hủy phiếu nháp (phiếu đã chốt không hủy được)."""
    denied = require_admin_api()
    if denied:
        return denied

    try:
        audit_id = clean_value((request.get_json(silent=True) or {}).get("id"))
        sheet, rows, header = read_audits(fresh=True)
        row = next((item for item in rows if item["Mã đánh giá"] == audit_id), None)
        if not row:
            return json_error("Không tìm thấy phiếu đánh giá này.", 404)
        if row["Trạng thái"] != AUDIT_DRAFT:
            return json_error("Chỉ hủy được phiếu nháp.")
        row["Trạng thái"] = AUDIT_CANCELLED
        row["Cập nhật lúc"] = now_text()
        update_record(sheet, AUDIT_SHEET, AUDIT_HEADERS, header, row)
        return jsonify({"success": True})

    except Exception as e:
        print("API AUDIT CANCEL ERROR:", e)
        return json_error(str(e), 500)


@app.route("/api/5s/actions")
def api_actions():
    denied = require_login_api()
    if denied:
        return denied

    try:
        actions = [a for a in build_actions() if in_lab_scope(a["lab"])]
        try:
            incidents = [
                {"id": i["id"], "name": i["target_name"], "status": i["status"], "description": i["description"][:80],
                 "zone": zone_of_target(i["target"])}
                for i in build_incidents() if i["status"] != INCIDENT_DONE and in_lab_scope(i["lab"])
            ]
        except Exception:
            incidents = []

        # Lỗi lặp lại khi bàn giao trong 14 ngày (gợi ý tạo việc)
        today = datetime.now().date()
        try:
            stats = handover_code_counts(DEFAULT_LAB, today - timedelta(days=REPEAT_WINDOW_DAYS), today)
            repeats = [{"code": code, "count": count} for code, count in
                       sorted(stats["any"].items(), key=lambda pair: -pair[1]) if count >= REPEAT_MIN]
        except Exception:
            repeats = []

        return jsonify({
            "success": True,
            "data": actions,
            "incidents": incidents,
            "repeats": repeats,
            "repeat_window": REPEAT_WINDOW_DAYS,
            "repeat_min": REPEAT_MIN
        })

    except Exception as e:
        print("API ACTIONS ERROR:", e)
        return json_error(str(e), 500)


@app.route("/api/5s/actions", methods=["POST"])
def api_action_add():
    """KTV tạo việc khắc phục: từ lỗi lặp lại khi bàn giao hoặc tự ghi nhận."""
    denied = require_admin_api()
    if denied:
        return denied

    try:
        payload = request.get_json(silent=True) or {}
        code = clean_value(payload.get("code")).upper()
        problem = clean_value(payload.get("problem"))[:300]
        todo = clean_value(payload.get("todo"))[:300]
        owner = clean_value(payload.get("owner")) or session.get("user", "")
        due = iso_date(payload.get("due"))
        source = clean_value(payload.get("source"))
        incident_id = clean_value(payload.get("incident_id"))

        if not CELL_CODE_PATTERN.match(code):
            return json_error("Chọn khu và S (mã lỗi dạng B-S2).")
        if not problem:
            return json_error("Ghi vấn đề cần khắc phục.")
        if not due:
            return json_error("Chọn hạn khắc phục.")
        if source not in (ACTION_SOURCE_REPEAT, ACTION_SOURCE_STAFF):
            source = ACTION_SOURCE_STAFF

        sheet, rows, header = read_table(ACTION_SHEET, ACTION_HEADERS, fresh=True)
        record = new_action_record(
            source, clean_value(payload.get("source_id"))[:60],
            normalize_lab_name(payload.get("lab")) or DEFAULT_LAB, code, problem,
            clean_value(payload.get("evidence"))[:300], todo, owner, due.strftime("%d/%m/%Y"), incident_id
        )
        append_records(sheet, ACTION_SHEET, ACTION_HEADERS, header, [record])
        return jsonify({"success": True, "id": record["Mã khắc phục"]})

    except Exception as e:
        print("API ACTION ADD ERROR:", e)
        return json_error(str(e), 500)


@app.route("/api/5s/actions/update", methods=["POST"])
def api_action_update():
    """Cập nhật việc khắc phục. Đã khắc phục: bắt buộc ghi kết quả; lưu người + ngày xác nhận (SOP-05 bước 8)."""
    denied = require_admin_api()
    if denied:
        return denied

    try:
        payload = request.get_json(silent=True) or {}
        action_id = clean_value(payload.get("id"))
        status = clean_value(payload.get("status"))
        result = clean_value(payload.get("result"))[:300]

        if status not in ACTION_STATUSES:
            return json_error("Trạng thái không hợp lệ.")
        if status == ACTION_DONE and not result:
            return json_error("Ghi kết quả khắc phục (đã làm gì, ảnh sau khắc phục) trước khi xác nhận.")

        sheet, rows, header = read_table(ACTION_SHEET, ACTION_HEADERS, fresh=True)
        record = next((row for row in rows if row["Mã khắc phục"] == action_id), None)
        if not record:
            return json_error("Không tìm thấy việc khắc phục này.", 404)

        due = iso_date(payload.get("due")) if payload.get("due") else iso_date(record["Hạn"])
        record.update({
            "Trạng thái": status,
            "Việc cần làm": clean_value(payload.get("todo"))[:300] if "todo" in payload else record["Việc cần làm"],
            "Người phụ trách": clean_value(payload.get("owner")) if "owner" in payload else record["Người phụ trách"],
            "Hạn": due.strftime("%d/%m/%Y") if due else record["Hạn"],
            "Mã sự cố": clean_value(payload.get("incident_id")) if "incident_id" in payload else record["Mã sự cố"],
            "Kết quả": result,
            "Người xác nhận": session.get("user", "") if status == ACTION_DONE else "",
            "Ngày xác nhận": now_text() if status == ACTION_DONE else "",
            "Cập nhật lúc": now_text(),
        })
        update_record(sheet, ACTION_SHEET, ACTION_HEADERS, header, record)
        return jsonify({"success": True})

    except Exception as e:
        print("API ACTION UPDATE ERROR:", e)
        return json_error(str(e), 500)


@app.route("/api/5s/summary")
def api_audit_summary():
    """Phiếu đã chốt cho trang Hiệu suất: xu hướng theo thời gian, theo S và khu, việc khắc phục."""
    denied = require_login_api()
    if denied:
        return denied

    try:
        _, cell_rows, _ = read_audit_cells()
        audits = []
        for row in reversed(finished_audits()):
            if not in_lab_scope(row["Phòng Lab"]):
                continue
            item = audit_public(row, cells_of(row["Mã đánh giá"], cell_rows))
            item.pop("low", None)
            audits.append(item)

        actions = [a for a in build_actions() if in_lab_scope(a["lab"])]
        return jsonify({
            "success": True,
            "audits": audits,
            "recent": audits[-4:],   # cũ -> mới
            "five_s": FIVE_S,
            "zones": [{"zone": z["zone"], "name": z["name"]} for z in AUDIT_ZONES],
            "actions": {
                "total": len(actions),
                "open": len([a for a in actions if a["status"] != ACTION_DONE]),
                "overdue": len([a for a in actions if a["overdue"]]),
                "done": len([a for a in actions if a["status"] == ACTION_DONE]),
            }
        })

    except Exception as e:
        print("API AUDIT SUMMARY ERROR:", e)
        return json_error(str(e), 500)


# =========================
# RUN
# =========================

if __name__ == "__main__":
    app.run(debug=True)
