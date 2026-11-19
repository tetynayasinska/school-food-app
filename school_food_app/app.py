import calendar
import datetime
import io
import os
import psycopg2
from psycopg2.extras import RealDictCursor

import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
import pandas as pd
from flask import (
    Flask,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "stavishche-lyceum-2-secret-key")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STUDENTS_XLSX = os.path.join(BASE_DIR, "students.xlsx")

# Змінна середовища для підключення до БД (Supabase / Neon / Render)
DATABASE_URL = os.environ.get("DATABASE_URL")

ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "jasinska33")

TEACHER_PASSWORDS = {
    "1-А": "0126", "1-Б": "0226", "2-А": "0326", "2-Б": "0426",
    "3-А": "0526", "3-Б": "0626", "4-А": "0726", "4-Б": "0826",
    "5-А": "0926", "5-Б": "1026", "6-А": "1126", "6-Б": "1226",
    "6-В": "1326", "7-А": "1426", "7-Б": "1526", "8-А": "1626",
    "8-Б": "1726", "9-А": "1826", "9-Б": "1926", "10-А": "2026",
    "10-Б": "2126", "10-В": "2226", "11-А": "2326", "11-Б": "2426",
}
PASSWORD_TO_CLASS = {v: k for k, v in TEACHER_PASSWORDS.items()}

UKR_ALPHABET = "абвгґдеєжзиіїйклмнопрстуфхцчшщьюя'’-"
UKR_ORDER = {ch: idx for idx, ch in enumerate(UKR_ALPHABET)}

MONTHS_GENITIVE = {
    1: "січні", 2: "лютому", 3: "березні", 4: "квітні", 5: "травні", 6: "червні",
    7: "липні", 8: "серпні", 9: "вересні", 10: "жовтні", 11: "листопаді", 12: "грудні",
}

MONTHS_NOMINATIVE = {
    1: "Січень", 2: "Лютий", 3: "Березень", 4: "Квітень", 5: "Травень", 6: "Червень",
    7: "Липень", 8: "Серпень", 9: "Вересень", 10: "Жовтень", 11: "Листопад", 12: "Грудень",
}

WEEKLY_MENU = {
    0: {"day": "Понеділок", "img": "1.jpg", "text": "Комплексний обід (Понеділок)"},
    1: {"day": "Вівторок", "img": "2.jpg", "text": "Комплексний обід (Вівторок)"},
    2: {"day": "Середа", "img": "3.jpg", "text": "Комплексний обід (Середа)"},
    3: {"day": "Четвер", "img": "4.jpg", "text": "Комплексний обід (Четвер)"},
    4: {"day": "П'ятниця", "img": "5.jpg", "text": "Комплексний обід (П'ятниця)"},
    5: {"day": "Субота", "img": "monday.jpg", "text": "Вихідний день"},
    6: {"day": "Неділя", "img": "monday.jpg", "text": "Вихідний день"},
}


def ukr_sort_key(name: str):
    s = (name or "").lower().strip()
    return [UKR_ORDER.get(ch, 100 + ord(ch)) for ch in s]


def class_sort_key(class_name: str):
    parts = (class_name or "").strip().split("-")
    grade = int(parts[0]) if parts and parts[0].isdigit() else 99
    letter = parts[1] if len(parts) > 1 else ""
    return (grade, ukr_sort_key(letter))


def get_two_month_cutoff(today: datetime.date | None = None) -> datetime.date:
    """Повертає 1-ше число попереднього місяця (усе, що раніше — видаляється)."""
    today = today or datetime.date.today()
    if today.month == 1:
        return datetime.date(today.year - 1, 12, 1)
    return datetime.date(today.year, today.month - 1, 1)


def get_available_report_months(today: datetime.date | None = None):
    """Повертає список із 2 останніх доступних місяців (поточний і попередній)."""
    today = today or datetime.date.today()
    prev = get_two_month_cutoff(today)
    months = [(today.year, today.month), (prev.year, prev.month)]
    return [
        {
            "value": f"{y:04d}-{m:02d}",
            "label": f"{MONTHS_NOMINATIVE[m]} {y}",
        }
        for y, m in months
    ]


def cleanup_old_attendance(cur=None):
    """Видаляє записи про харчування, старіші за 2 останні місяці."""
    if not DATABASE_URL:
        return
    cutoff_str = get_two_month_cutoff().strftime("%Y-%m-%d")
    if cur is not None:
        cur.execute("DELETE FROM Attendance WHERE date < %s", (cutoff_str,))
        return

    conn = get_db_connection()
    local_cur = conn.cursor()
    local_cur.execute("DELETE FROM Attendance WHERE date < %s", (cutoff_str,))
    local_cur.close()
    conn.close()


def get_db_connection():
    if not DATABASE_URL:
        raise RuntimeError("DATABASE_URL is not set!")
    conn = psycopg2.connect(DATABASE_URL)
    conn.autocommit = True
    return conn


def init_db():
    if not DATABASE_URL:
        print("Skipping DB init: DATABASE_URL not set.")
        return

    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS Students (
            id SERIAL PRIMARY KEY,
            full_name VARCHAR(255) NOT NULL,
            class_name VARCHAR(50) NOT NULL,
            is_privileged INTEGER NOT NULL DEFAULT 0,
            is_active INTEGER NOT NULL DEFAULT 1,
            added_manually INTEGER NOT NULL DEFAULT 0,
            UNIQUE(full_name, class_name)
        );
    """)

    cur.execute("ALTER TABLE Students ADD COLUMN IF NOT EXISTS added_manually INTEGER NOT NULL DEFAULT 0;")

    cur.execute("""
        CREATE TABLE IF NOT EXISTS Attendance (
            id SERIAL PRIMARY KEY,
            student_id INTEGER NOT NULL REFERENCES Students(id) ON DELETE CASCADE,
            date VARCHAR(20) NOT NULL,
            status VARCHAR(50) NOT NULL,
            timestamp VARCHAR(50),
            UNIQUE(student_id, date)
        );
    """)

    cur.execute("CREATE INDEX IF NOT EXISTS idx_attendance_date ON Attendance(date);")
    cleanup_old_attendance(cur)
    cur.close()
    conn.close()


def sync_students_from_excel(filepath: str = STUDENTS_XLSX):
    if not os.path.exists(filepath) or not DATABASE_URL:
        return 0

    df = pd.read_excel(filepath)
    df.columns = [str(c).strip() for c in df.columns]

    has_pilga_col = "Пільговик" in df.columns
    conn = get_db_connection()
    cur = conn.cursor()
    active_keys = set()
    synced = 0

    for _, row in df.iterrows():
        if pd.isna(row.get("ПІБ")) or pd.isna(row.get("Клас")):
            continue

        full_name = " ".join(str(row["ПІБ"]).strip().split())
        class_name = str(row["Клас"]).strip().replace("A", "А").replace("B", "В").replace("C", "С").replace("I", "І")

        is_privileged = 0
        if has_pilga_col and pd.notna(row["Пільговик"]):
            val = str(row["Пільговик"]).strip().upper()
            if val in ("П", "1", "TRUE", "ТАК", "+"):
                is_privileged = 1

        active_keys.add((full_name, class_name))

        cur.execute(
            """
            INSERT INTO Students (full_name, class_name, is_privileged, is_active, added_manually)
            VALUES (%s, %s, %s, 1, 0)
            ON CONFLICT (full_name, class_name) 
            DO UPDATE SET is_privileged = EXCLUDED.is_privileged, is_active = 1;
            """,
            (full_name, class_name, is_privileged)
        )
        synced += 1

    cur.execute("SELECT id, full_name, class_name, COALESCE(added_manually, 0) FROM Students")
    existing = cur.fetchall()

    for st in existing:
        # Не деактивуємо учнів, які були додані адміністратором вручну через сайт
        if (st[1], st[2]) not in active_keys and st[3] == 0:
            cur.execute("UPDATE Students SET is_active = 0 WHERE id = %s", (st[0],))

    cur.close()
    conn.close()
    return synced


def get_working_days(year: int, month: int):
    cal = calendar.Calendar(firstweekday=0)
    return [d for d in cal.itermonthdates(year, month) if d.month == month and d.weekday() < 5]


def build_monthly_report_excel(year: int, month: int) -> io.BytesIO:
    conn = get_db_connection()
    cur = conn.cursor(cursor_factory=RealDictCursor)

    cleanup_old_attendance(cur)

    cur.execute("SELECT id, full_name, class_name, is_privileged FROM Students WHERE is_active = 1")
    students = cur.fetchall()

    month_prefix = f"{year:04d}-{month:02d}-%"
    cur.execute("SELECT student_id, date, status FROM Attendance WHERE date LIKE %s", (month_prefix,))
    att_rows = cur.fetchall()

    cur.close()
    conn.close()

    attendance_map = {}
    for row in att_rows:
        st_val = str(row["status"]).strip()
        if st_val in ("Їсть", "1"):
            attendance_map[(row["student_id"], row["date"])] = 1
        elif st_val in ("Не їсть", "0", "н", "Н"):
            attendance_map[(row["student_id"], row["date"])] = "н"

    working_days = get_working_days(year, month)
    month_gen = MONTHS_GENITIVE.get(month, "")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    thin = Side(style="thin", color="000000")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    font_director = Font(name="Arial", size=18, bold=False)
    font_title = Font(name="Arial", size=18, bold=True)
    font_hdr = Font(name="Arial", size=18, bold=True)
    font_day_hdr = Font(name="Arial", size=20, bold=True)
    font_data = Font(name="Arial", size=18, bold=False)
    font_sub_num = Font(name="Arial", size=16, bold=False)
    font_pilga = Font(name="Arial", size=20, bold=True, color="FFFF0000")
    font_total = Font(name="Arial", size=18, bold=True)
    font_grand_sum = Font(name="Arial", size=20, bold=True)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    sheets_config = [
        {"title": "1-4 класи", "grades": [1, 2, 3, 4], "grand_label": "УСЬОГО 1-4 класи"},
        {"title": "5-11 класи", "grades": [5, 6, 7, 8, 9, 10, 11], "grand_label": "Всього 5-11 класи"},
    ]

    for cfg in sheets_config:
        ws = wb.create_sheet(title=cfg["title"])
        ws.views.sheetView[0].showGridLines = True

        col_widths = {"A": 11.5, "B": 6.5, "C": 7.0, "D": 42.5, "E": 15.0, "AC": 14.5, "AI": 13.0}
        for col_l, w in col_widths.items():
            ws.column_dimensions[col_l].width = w
        for c_idx in range(6, 29):
            ws.column_dimensions[get_column_letter(c_idx)].width = 8.0

        ws.merge_cells("A1:E2")
        ws["A1"] = "Затверджую\nДиректор «Ставищенський ліцей №2»\nГ.Поліщук"
        ws["A1"].font = font_director
        ws["A1"].alignment = align_center

        ws.merge_cells("F2:P3")
        ws["F2"] = f"Список учнів {cfg['title']} Ставищенського ліцею №2,\nякі харчувалися у {month_gen} {year} року"
        ws["F2"].font = font_title
        ws["F2"].alignment = align_center

        ws["V4"] = f"Харчування {cfg['title']}"
        ws["V4"].font = Font(name="Arial", size=14, bold=True)

        ws.row_dimensions[1].height = 31
        ws.row_dimensions[2].height = 64
        ws.row_dimensions[3].height = 34
        ws.row_dimensions[4].height = 25
        ws.row_dimensions[5].height = 30

        ws["A5"] = "№"
        ws["D5"] = "Прізвище та ім'я учня"
        ws["E5"] = "Клас"
        ws["AC5"] = "Всього"

        for idx, d in enumerate(working_days[:23]):
            col_idx = 6 + idx
            cell = ws.cell(5, col_idx, value=d.day)
            cell.font = font_day_hdr
            cell.alignment = align_center
            cell.border = border_all

        for c_idx in range(1, 30):
            cell = ws.cell(5, c_idx)
            if (cell.value is not None and c_idx < 6) or c_idx == 29:
                cell.font = font_hdr
                cell.alignment = align_center
            cell.border = border_all

        ws.print_title_rows = "$5:$5"

        current_row = 6
        sheet_seq_num = 1
        parallel_total_rows = []

        classes_dict = {}
        for st in students:
            cls = st["class_name"]
            grade_str = cls.split("-")[0]
            if grade_str.isdigit() and int(grade_str) in cfg["grades"]:
                classes_dict.setdefault(cls, []).append(dict(st))

        for grade in cfg["grades"]:
            grade_classes = sorted([c for c in classes_dict if int(c.split("-")[0]) == grade], key=class_sort_key)
            if not grade_classes:
                continue

            parallel_start_row = current_row

            for cls in grade_classes:
                cls_students = classes_dict[cls]
                regular_st = sorted([s for s in cls_students if not s["is_privileged"]], key=lambda x: ukr_sort_key(x["full_name"]))
                pilga_st = sorted([s for s in cls_students if s["is_privileged"]], key=lambda x: ukr_sort_key(x["full_name"]))
                ordered_students = regular_st + pilga_st

                for class_seq_num, st in enumerate(ordered_students, start=1):
                    ws.row_dimensions[current_row].height = 28
                    c_a = ws.cell(current_row, 1, value=sheet_seq_num)
                    c_a.font = font_data; c_a.alignment = align_center
                    c_b = ws.cell(current_row, 2, value=class_seq_num)
                    c_b.font = font_sub_num; c_b.alignment = align_center
                    c_c = ws.cell(current_row, 3, value="п" if st["is_privileged"] else None)
                    c_c.font = font_pilga; c_c.alignment = align_center
                    c_d = ws.cell(current_row, 4, value=st["full_name"])
                    c_d.font = font_data; c_d.alignment = align_left
                    c_e = ws.cell(current_row, 5, value=cls)
                    c_e.font = font_data; c_e.alignment = align_center

                    for idx, d in enumerate(working_days[:23]):
                        col_idx = 6 + idx
                        date_str = d.strftime("%Y-%m-%d")
                        cell_val = attendance_map.get((st["id"], date_str), None)
                        c_day = ws.cell(current_row, col_idx, value=cell_val)
                        c_day.font = font_data; c_day.alignment = align_center

                    for col_idx in range(6 + len(working_days[:23]), 29):
                        c_empty = ws.cell(current_row, col_idx, value=None)
                        c_empty.font = font_data; c_empty.alignment = align_center

                    c_total = ws.cell(current_row, 29, value=f"=COUNTIF(F{current_row}:AB{current_row},1)")
                    c_total.font = font_data; c_total.alignment = align_center

                    for col_idx in range(1, 30):
                        ws.cell(current_row, col_idx).border = border_all

                    sheet_seq_num += 1
                    current_row += 1

            parallel_end_row = current_row - 1
            ws.row_dimensions[current_row].height = 30
            label_d = "УСЬОГО 1 класи" if grade == 1 else "УСЬОГО"
            ws.cell(current_row, 4, value=label_d).alignment = align_left
            ws.cell(current_row, 4).font = font_total
            ws.cell(current_row, 5, value=f"{grade} класи").alignment = align_center
            ws.cell(current_row, 5).font = font_total

            for col_idx in range(6, 29):
                col_l = get_column_letter(col_idx)
                cell = ws.cell(current_row, col_idx, value=f"=COUNTIF({col_l}{parallel_start_row}:{col_l}{parallel_end_row},1)")
                cell.font = font_total; cell.alignment = align_center

            for col_idx in range(1, 30):
                ws.cell(current_row, col_idx).border = border_all

            c_ai = ws.cell(current_row, 35, value=f"=SUM(F{current_row}:AB{current_row})")
            c_ai.font = font_grand_sum; c_ai.alignment = align_center
            parallel_total_rows.append(current_row)
            current_row += 1

        ws.row_dimensions[current_row].height = 32
        ws.cell(current_row, 4, value=cfg["grand_label"]).font = font_total
        ws.cell(current_row, 4).alignment = align_left

        for col_idx in range(6, 29):
            col_l = get_column_letter(col_idx)
            refs = "+".join(f"{col_l}{r}" for r in reversed(parallel_total_rows))
            cell = ws.cell(current_row, col_idx, value=f"={refs}" if refs else 0)
            cell.font = font_total; cell.alignment = align_center

        for col_idx in range(1, 30):
            ws.cell(current_row, col_idx).border = border_all

        c_grand_ai = ws.cell(current_row, 35, value=f"=SUM(F{current_row}:AB{current_row})")
        c_grand_ai.font = font_grand_sum; c_grand_ai.alignment = align_center

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


@app.route("/")
def index():
    session.pop("admin_logged_in", None)
    session.pop("teacher_class", None)

    today_dt = datetime.date.today()
    today_str = today_dt.strftime("%Y-%m-%d")
    menu = WEEKLY_MENU.get(today_dt.weekday(), WEEKLY_MENU[0])

    if DATABASE_URL:
        conn = get_db_connection()
        cur = conn.cursor(cursor_factory=RealDictCursor)
        cur.execute("SELECT DISTINCT class_name FROM Students WHERE is_active = 1")
        rows = cur.fetchall()
        cur.close()
        conn.close()
        classes = sorted([r["class_name"] for r in rows], key=class_sort_key)
    else:
        classes = []

    return render_template("index.html", menu=menu, today=today_str, classes=classes)


@app.route("/get_students/<class_name>")
def get_students(class_name):
    conn = get_db_connection()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute("SELECT id, full_name FROM Students WHERE is_active = 1 AND class_name = %s", (class_name,))
    rows = cur.fetchall()
    cur.close()
    conn.close()

    sorted_rows = sorted(rows, key=lambda r: ukr_sort_key(r["full_name"]))
    return jsonify([{"id": r["id"], "name": r["full_name"]} for r in sorted_rows])


@app.route("/get_student_status/<int:student_id>")
def get_student_status(student_id):
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    conn = get_db_connection()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute("SELECT status, timestamp FROM Attendance WHERE student_id = %s AND date = %s", (student_id, today_str))
    row = cur.fetchone()
    cur.close()
    conn.close()

    if row:
        return jsonify({"status": row["status"], "timestamp": row["timestamp"] or ""})
    return jsonify({"status": "Не відмітився", "timestamp": ""})


@app.route("/submit_attendance", methods=["POST"])
def submit_attendance():
    data = request.get_json(silent=True) or {}
    student_id = data.get("student_id")
    status = data.get("status")

    if not student_id or status not in ("Їсть", "Не їсть"):
        return jsonify({"success": False, "error": "Некоректні дані"}), 400

    student_id = int(student_id)
    today_str = datetime.date.today().strftime("%Y-%m-%d")

    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO Attendance (student_id, date, status, timestamp)
        VALUES (%s, %s, %s, '')
        ON CONFLICT (student_id, date) 
        DO UPDATE SET status = EXCLUDED.status, timestamp = '';
        """,
        (student_id, today_str, status)
    )
    cleanup_old_attendance(cur)
    cur.close()
    conn.close()

    return jsonify({"success": True, "status": status, "timestamp": ""})


@app.route("/teacher", methods=["GET", "POST"])
def teacher():
    if not session.get("teacher_class"):
        error = None
        if request.method == "POST":
            entered_password = request.form.get("password", "").strip()
            if entered_password in PASSWORD_TO_CLASS:
                session["teacher_class"] = PASSWORD_TO_CLASS[entered_password]
                return redirect(url_for("teacher"))
            error = "Невірний пароль вчителя"
        return render_template("teacher.html", login_required=True, error=error)

    selected_class = session.get("teacher_class")
    today_str = datetime.date.today().strftime("%Y-%m-%d")

    conn = get_db_connection()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute(
        """
        SELECT
            s.id,
            s.full_name,
            s.is_privileged,
            COALESCE(a.status, 'Не відмітився') AS current_status
        FROM Students s
        LEFT JOIN Attendance a ON a.student_id = s.id AND a.date = %s
        WHERE s.is_active = 1 AND s.class_name = %s
        """,
        (today_str, selected_class)
    )
    st_rows = cur.fetchall()
    cur.close()
    conn.close()

    students = sorted([dict(r) for r in st_rows], key=lambda x: ukr_sort_key(x["full_name"]))

    return render_template(
        "teacher.html",
        login_required=False,
        today=today_str,
        classes=[selected_class],
        selected_class=selected_class,
        students=students,
    )


@app.route("/admin", methods=["GET", "POST"])
def admin():
    if not session.get("admin_logged_in"):
        error = None
        if request.method == "POST":
            if request.form.get("password") == ADMIN_PASSWORD:
                session["admin_logged_in"] = True
                return redirect(url_for("admin"))
            error = "Невірний пароль адміністратора"
        return render_template("admin.html", login_required=True, error=error)

    today_dt = datetime.date.today()
    today_str = today_dt.strftime("%Y-%m-%d")

    conn = get_db_connection()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    cleanup_old_attendance(cur)
    cur.execute(
        """
        SELECT
            s.class_name,
            COUNT(s.id) AS total,
            SUM(CASE WHEN a.status = 'Їсть' THEN 1 ELSE 0 END) AS eating,
            SUM(CASE WHEN a.status = 'Не їсть' THEN 1 ELSE 0 END) AS skip,
            SUM(CASE WHEN a.status IS NULL THEN 1 ELSE 0 END) AS unmarked
        FROM Students s
        LEFT JOIN Attendance a ON a.student_id = s.id AND a.date = %s
        WHERE s.is_active = 1
        GROUP BY s.class_name
        """,
        (today_str,)
    )
    rows = cur.fetchall()
    cur.close()
    conn.close()

    class_stats = [
        {
            "class_name": r["class_name"],
            "total": int(r["total"]) if r["total"] else 0,
            "eating": int(r["eating"]) if r["eating"] else 0,
            "skip": int(r["skip"]) if r["skip"] else 0,
            "unmarked": int(r["unmarked"]) if r["unmarked"] else 0,
        }
        for r in sorted(rows, key=lambda x: class_sort_key(x["class_name"]))
    ]

    total_students = sum(c["total"] for c in class_stats)
    total_eating = sum(c["eating"] for c in class_stats)
    total_skip = sum(c["skip"] for c in class_stats)
    total_unmarked = sum(c["unmarked"] for c in class_stats)
    all_classes = sorted(TEACHER_PASSWORDS.keys(), key=class_sort_key)

    return render_template(
        "admin.html",
        login_required=False,
        today=today_str,
        available_months=get_available_report_months(today_dt),
        all_classes=all_classes,
        total_students=total_students,
        total_eating=total_eating,
        total_skip=total_skip,
        total_unmarked=total_unmarked,
        class_stats=class_stats,
    )


@app.route("/admin/add_student", methods=["POST"])
def add_student():
    if not session.get("admin_logged_in"):
        return redirect(url_for("admin"))

    full_name = " ".join(request.form.get("full_name", "").strip().split())
    class_name = request.form.get("class_name", "").strip()
    is_privileged = 1 if request.form.get("is_privileged") == "1" else 0

    if len(full_name) < 3 or class_name not in TEACHER_PASSWORDS:
        flash("Вкажіть коректне прізвище та ім'я учня і виберіть клас зі списку.", "danger")
        return redirect(url_for("admin"))

    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO Students (full_name, class_name, is_privileged, is_active, added_manually)
        VALUES (%s, %s, %s, 1, 1)
        ON CONFLICT (full_name, class_name)
        DO UPDATE SET
            is_privileged = EXCLUDED.is_privileged,
            is_active = 1,
            added_manually = 1;
        """,
        (full_name, class_name, is_privileged)
    )
    cur.close()
    conn.close()

    pilga_text = " (пільговик)" if is_privileged else ""
    flash(f"Учня «{full_name}» ({class_name}{pilga_text}) успішно додано до списку!", "success")
    return redirect(url_for("admin"))


@app.route("/download_excel")
def download_excel():
    if not session.get("admin_logged_in"):
        return redirect(url_for("admin"))

    today = datetime.date.today()
    cutoff = get_two_month_cutoff(today)

    month_param = request.args.get("month")
    if month_param and "-" in month_param:
        try:
            year_s, month_s = month_param.split("-")
            year, month = int(year_s), int(month_s)
            if (year, month) < (cutoff.year, cutoff.month):
                year, month = today.year, today.month
        except ValueError:
            year, month = today.year, today.month
    else:
        year, month = today.year, today.month

    excel_stream = build_monthly_report_excel(year, month)
    month_name = MONTHS_NOMINATIVE.get(month, f"{month:02d}")
    filename = f"Харчування_1_11_звіт_{month_name}_{year}.xlsx"

    return send_file(
        excel_stream,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/logout")
def logout():
    session.pop("teacher_class", None)
    session.pop("admin_logged_in", None)
    return redirect(url_for("index"))


# Викликаємо ініціалізацію та синхронізацію одразу при запуску (працює і з Gunicorn, і локально)
if DATABASE_URL:
    try:
        init_db()
        sync_students_from_excel(STUDENTS_XLSX)
    except Exception as e:
        print(f"DB Init/Sync Error: {e}")

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))