from flask import Flask, render_template, jsonify
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

app = Flask(__name__)

# =========================
# GOOGLE SHEET CONNECT
# =========================

scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

creds = Credentials.from_service_account_file(
    "credentials.json.json",
    scopes=scope
)

client = gspread.authorize(creds)

# =========================
# PAGE
# =========================

@app.route("/")
def dashboard():
    return render_template("dashboard.html")

# =========================
# API
# =========================

@app.route("/api/bookings")
def get_bookings():

    try:

        sheet = client.open(
            "PHIẾU ĐĂNG KÝ MƯỢN PHÒNG LAB – KHOA KỸ THUẬT VÀ QUẢN LÝ CÔNG NGHIỆP  (Câu trả lời)"
        ).worksheet("LAB_BOOKING")

        rows = sheet.get_all_records()

        bookings = []

        for row in rows:

            status = str(
                row.get("TRẠNG THÁI", "")
            ).strip()

            # CHỈ LẤY LỊCH ĐÃ XÁC NHẬN
            if status != "ĐÃ_XÁC_NHẬN":
                continue

            raw_date = str(
                row.get("8. Ngày bạn đăng ký sử dụng phòng là:", "")
            ).strip()

            if raw_date == "":
                continue

            try:

                formatted_date = datetime.strptime(
                    raw_date,
                    "%d/%m/%Y"
                ).strftime("%Y-%m-%d")

            except:
                continue

            bookings.append({
                "name": str(
                    row.get("1. Họ và tên của bạn là: ", "")
                ).strip(),

                "lab": str(
                    row.get("6. Phòng lab được lựa chọn là:", "")
                ).strip(),

                "shift": str(
                    row.get("9. Ca bạn đăng ký sử dụng là: ", "")
                ).strip(),

                "date": formatted_date,

                "status": status,

                "idKey": str(
                    row.get("ID_KEY", "")
                ).strip()
            })

        return jsonify(bookings)

    except Exception as e:

        return jsonify({
            "error": str(e)
        })

# =========================
# RUN
# =========================

if __name__ == "__main__":
    app.run(debug=True)

