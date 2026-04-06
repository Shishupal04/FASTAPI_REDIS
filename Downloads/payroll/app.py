import os
import shutil
from datetime import datetime
from functools import wraps

from flask import (Flask, flash, redirect, render_template,
                   request, send_file, session)
from openpyxl import Workbook
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

from database import init_db, DB_PATH


# ──────────────────────────────────────────────────────────────
# App setup
# ──────────────────────────────────────────────────────────────
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-change-in-prod")

# Use /tmp for all file writes on Render (only /tmp is writable on free tier)
EXPORT_DIR = "/tmp/payroll_exports"
BACKUP_DIR = "/tmp/payroll_backups"
os.makedirs(EXPORT_DIR, exist_ok=True)
os.makedirs(BACKUP_DIR, exist_ok=True)

# ── CRITICAL: initialise DB at module load so Gunicorn picks it up ──
init_db()


# ──────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────
def get_db():
    import sqlite3
    return sqlite3.connect(DB_PATH)


def role_required(allowed_roles):
    def decorator(f):
        @wraps(f)
        def wrapper(*args, **kwargs):
            if "user" not in session:
                return redirect("/login")
            if session.get("role") not in allowed_roles:
                return render_template("403.html"), 403
            return f(*args, **kwargs)
        return wrapper
    return decorator


# ──────────────────────────────────────────────────────────────
# Dashboard
# ──────────────────────────────────────────────────────────────
@app.route("/")
def dashboard():
    if "user" not in session:
        return redirect("/login")

    conn   = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM employees")
    emp_count = cursor.fetchone()[0]

    cursor.execute("SELECT SUM(net) FROM payroll")
    payout = cursor.fetchone()[0] or 0

    cursor.execute("SELECT COUNT(*) FROM payroll")
    processed = cursor.fetchone()[0]

    cursor.execute("SELECT * FROM employees ORDER BY id DESC LIMIT 5")
    employees = cursor.fetchall()
    conn.close()

    today = datetime.now().strftime("%A, %d %B %Y")
    return render_template("dashboard.html",
                           emp_count=emp_count, payout=payout,
                           processed=processed, employees=employees, today=today)


# ──────────────────────────────────────────────────────────────
# Auth
# ──────────────────────────────────────────────────────────────
@app.route("/login", methods=["GET", "POST"])
def login():
    if "user" in session:
        return redirect("/")

    error = None
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]

        conn   = get_db()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM users WHERE username=? AND password=?",
                       (username, password))
        user = cursor.fetchone()
        conn.close()

        if user:
            session["user"] = user[1]
            session["role"] = user[3]
            return redirect("/")
        error = "Invalid username or password."

    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")


# ──────────────────────────────────────────────────────────────
# Employees
# ──────────────────────────────────────────────────────────────
@app.route("/employees")
@role_required(["admin", "hr"])
def employees():
    conn   = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM employees ORDER BY id DESC")
    employees_list = cursor.fetchall()
    cursor.execute("SELECT id, name FROM companies")
    companies = cursor.fetchall()
    conn.close()
    return render_template("employees.html",
                           employees=employees_list, companies=companies)


@app.route("/add_employee", methods=["POST"])
@role_required(["admin", "hr"])
def add_employee():
    data = (
        request.form["company_id"], request.form["emp_code"],
        request.form["name"],       request.form["department"],
        request.form["designation"],request.form["doj"],
        request.form["basic"],      request.form["hra"],
        request.form["allowance"],  request.form["bank"],
        request.form["ifsc"],       request.form["pan"],
        request.form["uan"]
    )
    conn   = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO employees
            (company_id, emp_code, name, department, designation, doj,
             basic, hra, allowance, bank_account, ifsc, pan, uan)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
    """, data)
    conn.commit()
    conn.close()
    flash("Employee added successfully.", "success")
    return redirect("/employees")


# ──────────────────────────────────────────────────────────────
# Attendance
# ──────────────────────────────────────────────────────────────
@app.route("/attendance")
@role_required(["admin", "hr"])
def attendance():
    conn   = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT id, name FROM employees ORDER BY name")
    employees = cursor.fetchall()
    cursor.execute("""
        SELECT attendance.*, employees.name
        FROM attendance
        JOIN employees ON employees.id = attendance.emp_id
        ORDER BY attendance.month DESC
    """)
    records = cursor.fetchall()
    conn.close()
    return render_template("attendance.html", employees=employees, records=records)


@app.route("/add_attendance", methods=["POST"])
@role_required(["admin", "hr"])
def add_attendance():
    data = (request.form["emp_id"], request.form["month"],
            request.form["working"], request.form["present"],
            request.form["lop"])
    conn   = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO attendance (emp_id, month, working_days, present_days, lop_days)
        VALUES (?,?,?,?,?)
    """, data)
    conn.commit()
    conn.close()
    flash("Attendance saved.", "success")
    return redirect("/attendance")


# ──────────────────────────────────────────────────────────────
# Payroll
# ──────────────────────────────────────────────────────────────
@app.route("/payroll")
@role_required(["admin", "accountant"])
def payroll():
    conn   = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT id, name FROM employees ORDER BY name")
    employees = cursor.fetchall()
    cursor.execute("""
        SELECT payroll.*, employees.name
        FROM payroll
        JOIN employees ON employees.id = payroll.emp_id
        ORDER BY payroll.month DESC
    """)
    records = cursor.fetchall()
    conn.close()
    return render_template("payroll.html", employees=employees, records=records)


@app.route("/process_payroll", methods=["POST"])
@role_required(["admin", "accountant"])
def process_payroll():
    emp_id = request.form["emp_id"]
    month  = request.form["month"]

    conn   = get_db()
    cursor = conn.cursor()

    # Prevent duplicate
    cursor.execute("SELECT id FROM payroll WHERE emp_id=? AND month=?",
                   (emp_id, month))
    if cursor.fetchone():
        conn.close()
        flash("Payroll already processed for this employee and month.", "warning")
        return redirect("/payroll")

    # Base salary
    cursor.execute("SELECT basic, hra, allowance FROM employees WHERE id=?", (emp_id,))
    emp = cursor.fetchone()
    if not emp:
        conn.close()
        flash("Employee not found.", "danger")
        return redirect("/payroll")

    basic, hra, allowance = (float(x or 0) for x in emp)

    # Latest salary revision
    cursor.execute("""
        SELECT basic, hra, allowances FROM salary_history
        WHERE emp_id=? AND effective_date<=?
        ORDER BY effective_date DESC LIMIT 1
    """, (emp_id, month + "-01"))
    revision = cursor.fetchone()
    if revision:
        basic, hra, allowance = (float(x or 0) for x in revision)

    gross = basic + hra + allowance

    # Attendance
    cursor.execute("SELECT working_days, present_days FROM attendance WHERE emp_id=? AND month=?",
                   (emp_id, month))
    att = cursor.fetchone()
    if not att:
        conn.close()
        flash("Attendance not found. Please mark attendance first.", "danger")
        return redirect("/payroll")

    working, present = int(att[0] or 26), int(att[1] or 0)
    lop_days         = max(0, working - present)
    per_day          = gross / working if working > 0 else 0
    lop_amount       = per_day * lop_days
    gross_after_lop  = gross - lop_amount

    # Overtime
    cursor.execute("""
        SELECT SUM(hours) FROM overtime
        WHERE emp_id=? AND strftime('%Y-%m', ot_date)=?
    """, (emp_id, month))
    ot_hours    = float(cursor.fetchone()[0] or 0)
    hourly_rate = (gross / 26 / 8) if gross > 0 else 0
    ot_amount   = ot_hours * hourly_rate * 1.5

    # Deductions
    pf  = basic * 0.12
    esi = gross_after_lop * 0.0075 if gross_after_lop <= 21000 else 0
    pt  = 0 if gross_after_lop <= 15000 else (150 if gross_after_lop <= 20000 else 200)
    tds = gross_after_lop * 0.05
    net = gross_after_lop + ot_amount - (pf + esi + pt + tds)

    cursor.execute("""
        INSERT INTO payroll
            (emp_id, month, basic, hra, allowance, gross,
             pf, esi, pt, tds, lop, net)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
    """, (
        emp_id, month,
        round(basic, 2), round(hra, 2), round(allowance, 2), round(gross, 2),
        round(pf, 2), round(esi, 2), round(pt, 2), round(tds, 2),
        round(lop_amount, 2), round(net, 2)
    ))
    conn.commit()
    conn.close()
    flash(f"Payroll processed successfully. Net salary: ₹{net:,.2f}", "success")
    return redirect("/payroll")


# ──────────────────────────────────────────────────────────────
# Payslip PDF
# ──────────────────────────────────────────────────────────────
@app.route("/payslip/<int:pay_id>")
@role_required(["admin", "accountant", "hr"])
def payslip(pay_id):
    conn   = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT payroll.*, employees.name, employees.designation,
               employees.department, employees.emp_code
        FROM payroll
        JOIN employees ON employees.id = payroll.emp_id
        WHERE payroll.id=?
    """, (pay_id,))
    data = cursor.fetchone()
    conn.close()

    if not data:
        return "Payslip not found.", 404

    file_path = os.path.join(EXPORT_DIR, f"payslip_{pay_id}.pdf")
    c = canvas.Canvas(file_path, pagesize=A4)
    W, H = A4

    # Header bar
    c.setFillColorRGB(0.07, 0.15, 0.40)
    c.rect(0, H - 80, W, 80, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.setFont("Helvetica-Bold", 18)
    c.drawString(40, H - 40, "SALARY PAYSLIP")
    c.setFont("Helvetica", 10)
    c.drawString(40, H - 60, f"Month: {data[2]}")

    # Employee details
    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, H - 110, "Employee Details")
    c.line(40, H - 115, W - 40, H - 115)
    c.setFont("Helvetica", 10)
    c.drawString(40,  H - 135, f"Name        : {data[13]}")
    c.drawString(40,  H - 150, f"Emp Code    : {data[16]}")
    c.drawString(300, H - 135, f"Department  : {data[15]}")
    c.drawString(300, H - 150, f"Designation : {data[14]}")

    # Earnings
    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, H - 185, "Earnings")
    c.line(40, H - 190, 260, H - 190)
    c.setFont("Helvetica", 10)
    c.drawString(40, H - 210, f"Basic Salary   : Rs {data[3]:>10,.2f}")
    c.drawString(40, H - 225, f"HRA            : Rs {data[4]:>10,.2f}")
    c.drawString(40, H - 240, f"Allowances     : Rs {data[5]:>10,.2f}")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(40, H - 260, f"Gross Salary   : Rs {data[6]:>10,.2f}")

    # Deductions
    c.setFont("Helvetica-Bold", 11)
    c.drawString(310, H - 185, "Deductions")
    c.line(310, H - 190, W - 40, H - 190)
    c.setFont("Helvetica", 10)
    c.drawString(310, H - 210, f"Provident Fund : Rs {data[7]:>10,.2f}")
    c.drawString(310, H - 225, f"ESI            : Rs {data[8]:>10,.2f}")
    c.drawString(310, H - 240, f"Prof. Tax (PT) : Rs {data[9]:>10,.2f}")
    c.drawString(310, H - 255, f"TDS            : Rs {data[10]:>10,.2f}")
    c.drawString(310, H - 270, f"LOP Deduction  : Rs {data[11]:>10,.2f}")

    # Net salary banner
    c.setFillColorRGB(0.07, 0.15, 0.40)
    c.rect(40, H - 315, W - 80, 30, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50,    H - 305, "NET SALARY")
    c.drawRightString(W - 50, H - 305, f"Rs {data[12]:,.2f}")

    c.setFillColorRGB(0.5, 0.5, 0.5)
    c.setFont("Helvetica", 8)
    c.drawCentredString(W / 2, 30,
        "This is a computer-generated payslip and does not require a signature.")
    c.save()

    return send_file(file_path, as_attachment=True,
                     download_name=f"Payslip_{data[13]}_{data[2]}.pdf")


# ──────────────────────────────────────────────────────────────
# Reports
# ──────────────────────────────────────────────────────────────
@app.route("/reports")
@role_required(["admin", "accountant"])
def reports():
    return render_template("reports.html")


@app.route("/salary_register")
@role_required(["admin", "accountant"])
def salary_register():
    conn   = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT employees.name, payroll.month, payroll.gross,
               payroll.pf, payroll.esi, payroll.pt, payroll.tds, payroll.net
        FROM payroll JOIN employees ON employees.id = payroll.emp_id
        ORDER BY payroll.month DESC
    """)
    data = cursor.fetchall()
    conn.close()
    return render_template("salary_register.html", data=data)


@app.route("/bank_sheet")
@role_required(["admin", "accountant"])
def bank_sheet():
    conn   = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT employees.name, employees.bank_account,
               employees.ifsc, payroll.net
        FROM payroll JOIN employees ON employees.id = payroll.emp_id
    """)
    data = cursor.fetchall()
    conn.close()
    return render_template("bank_sheet.html", data=data)


@app.route("/statutory_report")
@role_required(["admin", "accountant"])
def statutory_report():
    conn   = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT employees.name, payroll.pf, payroll.esi
        FROM payroll JOIN employees ON employees.id = payroll.emp_id
    """)
    data = cursor.fetchall()
    conn.close()
    return render_template("statutory_report.html", data=data)


# ──────────────────────────────────────────────────────────────
# Excel Exports
# ──────────────────────────────────────────────────────────────
@app.route("/export_salary_register")
@role_required(["admin", "accountant"])
def export_salary_register():
    conn   = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT employees.name, payroll.month, payroll.gross,
               payroll.pf, payroll.esi, payroll.pt, payroll.tds, payroll.net
        FROM payroll JOIN employees ON employees.id = payroll.emp_id
    """)
    data = cursor.fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Salary Register"
    ws.append(["Name", "Month", "Gross", "PF", "ESI", "PT", "TDS", "Net"])
    for row in data:
        ws.append(list(row))

    path = os.path.join(EXPORT_DIR, "salary_register.xlsx")
    wb.save(path)
    return send_file(path, as_attachment=True,
                     download_name="Salary_Register.xlsx")


@app.route("/export_bank_sheet")
@role_required(["admin", "accountant"])
def export_bank_sheet():
    conn   = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT employees.name, employees.bank_account,
               employees.ifsc, payroll.net
        FROM payroll JOIN employees ON employees.id = payroll.emp_id
    """)
    data = cursor.fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Bank Sheet"
    ws.append(["Name", "Account", "IFSC", "Amount"])
    for row in data:
        ws.append(list(row))

    path = os.path.join(EXPORT_DIR, "bank_sheet.xlsx")
    wb.save(path)
    return send_file(path, as_attachment=True,
                     download_name="Bank_Sheet.xlsx")


# ──────────────────────────────────────────────────────────────
# Leave
# ──────────────────────────────────────────────────────────────
@app.route("/leave", methods=["GET", "POST"])
@role_required(["admin", "hr"])
def leave():
    conn   = get_db()
    cursor = conn.cursor()

    if request.method == "POST":
        cursor.execute("""
            INSERT INTO leaves (emp_id, leave_date, leave_type)
            VALUES (?,?,?)
        """, (request.form["emp_id"], request.form["leave_date"],
              request.form["leave_type"]))
        conn.commit()
        conn.close()
        flash("Leave recorded.", "success")
        return redirect("/leave")

    cursor.execute("SELECT id, name FROM employees ORDER BY name")
    employees = cursor.fetchall()
    cursor.execute("""
        SELECT leaves.*, employees.name FROM leaves
        JOIN employees ON employees.id = leaves.emp_id
        ORDER BY leaves.leave_date DESC
    """)
    records = cursor.fetchall()
    conn.close()
    return render_template("leave.html", employees=employees, records=records)


# ──────────────────────────────────────────────────────────────
# Overtime
# ──────────────────────────────────────────────────────────────
@app.route("/overtime", methods=["GET", "POST"])
@role_required(["admin", "hr"])
def overtime():
    conn   = get_db()
    cursor = conn.cursor()

    if request.method == "POST":
        cursor.execute("""
            INSERT INTO overtime (emp_id, ot_date, hours)
            VALUES (?,?,?)
        """, (request.form["emp_id"], request.form["ot_date"],
              request.form["hours"]))
        conn.commit()
        conn.close()
        flash("Overtime recorded.", "success")
        return redirect("/overtime")

    cursor.execute("SELECT id, name FROM employees ORDER BY name")
    employees = cursor.fetchall()
    cursor.execute("""
        SELECT overtime.*, employees.name FROM overtime
        JOIN employees ON employees.id = overtime.emp_id
        ORDER BY overtime.ot_date DESC
    """)
    records = cursor.fetchall()
    conn.close()
    return render_template("overtime.html", employees=employees, records=records)


# ──────────────────────────────────────────────────────────────
# Salary Revision
# ──────────────────────────────────────────────────────────────
@app.route("/salary_revision", methods=["GET", "POST"])
@role_required(["admin", "hr"])
def salary_revision():
    conn   = get_db()
    cursor = conn.cursor()

    if request.method == "POST":
        cursor.execute("""
            INSERT INTO salary_history (emp_id, effective_date, basic, hra, allowances)
            VALUES (?,?,?,?,?)
        """, (request.form["emp_id"], request.form["date"],
              request.form["basic"], request.form["hra"],
              request.form["allowances"]))
        conn.commit()
        conn.close()
        flash("Salary revision saved.", "success")
        return redirect("/salary_revision")

    cursor.execute("SELECT id, name FROM employees ORDER BY name")
    employees = cursor.fetchall()
    conn.close()
    return render_template("salary_revision.html", employees=employees)


@app.route("/salary_history/<int:emp_id>")
@role_required(["admin", "hr", "accountant"])
def salary_history(emp_id):
    conn   = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT effective_date, basic, hra, allowances
        FROM salary_history WHERE emp_id=?
        ORDER BY effective_date DESC
    """, (emp_id,))
    records = cursor.fetchall()
    conn.close()
    return render_template("salary_history.html", records=records)


# ──────────────────────────────────────────────────────────────
# Backup / Restore
# ──────────────────────────────────────────────────────────────
@app.route("/backup")
@role_required(["admin"])
def backup():
    timestamp   = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = os.path.join(BACKUP_DIR, f"payroll_{timestamp}.db")
    shutil.copy(DB_PATH, backup_file)
    flash(f"Backup created: payroll_{timestamp}.db", "success")
    return redirect("/backup_page")


@app.route("/restore/<filename>")
@role_required(["admin"])
def restore(filename):
    safe_name   = os.path.basename(filename)
    backup_file = os.path.join(BACKUP_DIR, safe_name)
    if not os.path.exists(backup_file):
        flash("Backup file not found.", "danger")
        return redirect("/backup_page")
    shutil.copy(backup_file, DB_PATH)
    flash("Database restored successfully.", "success")
    return redirect("/backup_page")


@app.route("/backup_page")
@role_required(["admin"])
def backup_page():
    files = sorted(os.listdir(BACKUP_DIR), reverse=True) if os.path.exists(BACKUP_DIR) else []
    return render_template("backup.html", files=files)


# ──────────────────────────────────────────────────────────────
# Companies
# ──────────────────────────────────────────────────────────────
@app.route("/companies")
@role_required(["admin"])
def companies():
    conn   = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM companies")
    data = cursor.fetchall()
    conn.close()
    return render_template("companies.html", companies=data)


@app.route("/add_company", methods=["POST"])
@role_required(["admin"])
def add_company():
    conn   = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO companies (name, address, pf_number, esi_number)
        VALUES (?,?,?,?)
    """, (request.form["name"], request.form["address"],
          request.form["pf"],   request.form["esi"]))
    conn.commit()
    conn.close()
    flash("Company added.", "success")
    return redirect("/companies")


# ──────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app.run(debug=True)
