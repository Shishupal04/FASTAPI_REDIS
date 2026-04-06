import os
import shutil
import traceback
from datetime import datetime
from functools import wraps

from flask import (Flask, flash, redirect, render_template,
                   request, send_file, session)
from openpyxl import Workbook
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

from database import init_db, get_db

# ──────────────────────────────────────────────────────────────
# App bootstrap
# ──────────────────────────────────────────────────────────────
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "payroll-super-secret-2024-xK9mP")

EXPORT_DIR = "/tmp/payroll_exports"
BACKUP_DIR = "/tmp/payroll_backups"
os.makedirs(EXPORT_DIR, exist_ok=True)
os.makedirs(BACKUP_DIR, exist_ok=True)

# Initialise DB at module load — runs for BOTH gunicorn AND python app.py
init_db()


# ──────────────────────────────────────────────────────────────
# Global error handler — shows real error instead of blank 500
# ──────────────────────────────────────────────────────────────
@app.errorhandler(500)
def internal_error(error):
    tb = traceback.format_exc()
    print("[500 ERROR]", tb)
    return f"""
    <html><body style='font-family:monospace;padding:30px;background:#1e293b;color:#f8fafc'>
    <h2 style='color:#ef4444'>Internal Server Error</h2>
    <p>Please copy this and share with support:</p>
    <pre style='background:#0f172a;padding:20px;border-radius:8px;
                overflow:auto;font-size:12px'>{tb}</pre>
    <a href='/' style='color:#60a5fa'>← Go Home</a>
    </body></html>
    """, 500


# ──────────────────────────────────────────────────────────────
# Decorators
# ──────────────────────────────────────────────────────────────
def login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if "user" not in session:
            return redirect("/login")
        return f(*args, **kwargs)
    return wrapper


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
# Health check — visit /health to confirm app + DB are working
# ──────────────────────────────────────────────────────────────
@app.route("/health")
def health():
    try:
        conn   = get_db()
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM users")
        user_count = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM employees")
        emp_count = cursor.fetchone()[0]
        conn.close()
        from database import DB_PATH
        return {
            "status": "ok",
            "db_path": DB_PATH,
            "users": user_count,
            "employees": emp_count,
        }, 200
    except Exception as e:
        return {"status": "error", "detail": str(e)}, 500


# ──────────────────────────────────────────────────────────────
# Dashboard
# ──────────────────────────────────────────────────────────────
@app.route("/")
@login_required
def dashboard():
    conn   = get_db()
    cursor = conn.cursor()

    cursor.execute("SELECT COUNT(*) FROM employees")
    emp_count = cursor.fetchone()[0]

    cursor.execute("SELECT COALESCE(SUM(net),0) FROM payroll")
    payout = cursor.fetchone()[0]

    cursor.execute("SELECT COUNT(*) FROM payroll")
    processed = cursor.fetchone()[0]

    cursor.execute("SELECT * FROM employees ORDER BY id DESC LIMIT 5")
    employees = cursor.fetchall()

    conn.close()
    today = datetime.now().strftime("%A, %d %B %Y")

    return render_template("dashboard.html",
                           emp_count=emp_count,
                           payout=float(payout),
                           processed=processed,
                           employees=employees,
                           today=today)


# ──────────────────────────────────────────────────────────────
# Auth
# ──────────────────────────────────────────────────────────────
@app.route("/login", methods=["GET", "POST"])
def login():
    if "user" in session:
        return redirect("/")

    error = None
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()

        conn   = get_db()
        cursor = conn.cursor()
        cursor.execute(
            "SELECT * FROM users WHERE username=? AND password=?",
            (username, password)
        )
        user = cursor.fetchone()
        conn.close()

        if user:
            session.permanent = True
            session["user"] = user["username"]
            session["role"] = user["role"]
            flash(f"Welcome back, {user['username']}!", "success")
            return redirect("/")
        error = "Invalid username or password. Please try again."

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
                           employees=employees_list,
                           companies=companies)


@app.route("/add_employee", methods=["POST"])
@role_required(["admin", "hr"])
def add_employee():
    conn   = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO employees
            (company_id, emp_code, name, department, designation, doj,
             basic, hra, allowance, bank_account, ifsc, pan, uan)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
    """, (
        request.form.get("company_id", 1),
        request.form.get("emp_code", ""),
        request.form.get("name", ""),
        request.form.get("department", ""),
        request.form.get("designation", ""),
        request.form.get("doj", ""),
        float(request.form.get("basic", 0) or 0),
        float(request.form.get("hra", 0) or 0),
        float(request.form.get("allowance", 0) or 0),
        request.form.get("bank", ""),
        request.form.get("ifsc", ""),
        request.form.get("pan", ""),
        request.form.get("uan", ""),
    ))
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
        SELECT a.*, e.name
        FROM attendance a
        JOIN employees e ON e.id = a.emp_id
        ORDER BY a.month DESC
    """)
    records = cursor.fetchall()
    conn.close()
    return render_template("attendance.html",
                           employees=employees, records=records)


@app.route("/add_attendance", methods=["POST"])
@role_required(["admin", "hr"])
def add_attendance():
    conn   = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO attendance (emp_id, month, working_days, present_days, lop_days)
        VALUES (?,?,?,?,?)
    """, (
        request.form["emp_id"],
        request.form["month"],
        int(request.form.get("working", 26)),
        int(request.form.get("present", 0)),
        int(request.form.get("lop", 0)),
    ))
    conn.commit()
    conn.close()
    flash("Attendance saved successfully.", "success")
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
        SELECT p.*, e.name
        FROM payroll p
        JOIN employees e ON e.id = p.emp_id
        ORDER BY p.month DESC
    """)
    records = cursor.fetchall()
    conn.close()
    return render_template("payroll.html",
                           employees=employees, records=records)


@app.route("/process_payroll", methods=["POST"])
@role_required(["admin", "accountant"])
def process_payroll():
    emp_id = request.form["emp_id"]
    month  = request.form["month"]

    conn   = get_db()
    cursor = conn.cursor()

    # Prevent duplicate
    cursor.execute(
        "SELECT id FROM payroll WHERE emp_id=? AND month=?", (emp_id, month))
    if cursor.fetchone():
        conn.close()
        flash("Payroll already processed for this employee and month.", "warning")
        return redirect("/payroll")

    # Base salary
    cursor.execute(
        "SELECT basic, hra, allowance FROM employees WHERE id=?", (emp_id,))
    emp = cursor.fetchone()
    if not emp:
        conn.close()
        flash("Employee not found.", "danger")
        return redirect("/payroll")

    basic     = float(emp["basic"]     or 0)
    hra       = float(emp["hra"]       or 0)
    allowance = float(emp["allowance"] or 0)

    # Apply latest salary revision
    cursor.execute("""
        SELECT basic, hra, allowances FROM salary_history
        WHERE emp_id=? AND effective_date<=?
        ORDER BY effective_date DESC LIMIT 1
    """, (emp_id, month + "-01"))
    rev = cursor.fetchone()
    if rev:
        basic     = float(rev["basic"]      or 0)
        hra       = float(rev["hra"]        or 0)
        allowance = float(rev["allowances"] or 0)

    gross = basic + hra + allowance

    # Attendance
    cursor.execute(
        "SELECT working_days, present_days FROM attendance WHERE emp_id=? AND month=?",
        (emp_id, month))
    att = cursor.fetchone()
    if not att:
        conn.close()
        flash("Attendance not found. Please mark attendance first.", "danger")
        return redirect("/payroll")

    working         = int(att["working_days"] or 26)
    present         = int(att["present_days"] or 0)
    lop_days        = max(0, working - present)
    per_day         = gross / working if working > 0 else 0
    lop_amount      = per_day * lop_days
    gross_after_lop = gross - lop_amount

    # Overtime
    cursor.execute("""
        SELECT COALESCE(SUM(hours), 0) FROM overtime
        WHERE emp_id=? AND strftime('%Y-%m', ot_date)=?
    """, (emp_id, month))
    ot_hours    = float(cursor.fetchone()[0])
    hourly_rate = (gross / 26 / 8) if gross > 0 else 0
    ot_amount   = ot_hours * hourly_rate * 1.5

    # Deductions
    pf  = basic * 0.12
    esi = gross_after_lop * 0.0075 if gross_after_lop <= 21000 else 0
    pt  = (0 if gross_after_lop <= 15000
           else 150 if gross_after_lop <= 20000
           else 200)
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
        round(lop_amount, 2), round(net, 2),
    ))
    conn.commit()
    conn.close()
    flash(f"Payroll processed. Net salary: ₹{net:,.2f}", "success")
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
        SELECT p.*, e.name, e.designation, e.department, e.emp_code
        FROM payroll p
        JOIN employees e ON e.id = p.emp_id
        WHERE p.id=?
    """, (pay_id,))
    row = cursor.fetchone()
    conn.close()

    if not row:
        flash("Payslip not found.", "danger")
        return redirect("/payroll")

    d = dict(row)
    file_path = os.path.join(EXPORT_DIR, f"payslip_{pay_id}.pdf")

    c = canvas.Canvas(file_path, pagesize=A4)
    W, H = A4

    # Header
    c.setFillColorRGB(0.07, 0.15, 0.40)
    c.rect(0, H - 80, W, 80, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.setFont("Helvetica-Bold", 18)
    c.drawString(40, H - 40, "SALARY PAYSLIP")
    c.setFont("Helvetica", 10)
    c.drawString(40, H - 60, f"Month: {d['month']}")

    # Employee info
    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, H - 110, "Employee Details")
    c.line(40, H - 115, W - 40, H - 115)
    c.setFont("Helvetica", 10)
    c.drawString(40,  H - 135, f"Name        : {d['name']}")
    c.drawString(40,  H - 150, f"Emp Code    : {d['emp_code']}")
    c.drawString(300, H - 135, f"Department  : {d['department']}")
    c.drawString(300, H - 150, f"Designation : {d['designation']}")

    # Earnings
    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, H - 185, "Earnings")
    c.line(40, H - 190, 260, H - 190)
    c.setFont("Helvetica", 10)
    c.drawString(40, H - 210, f"Basic Salary   :  Rs {d['basic']:>10,.2f}")
    c.drawString(40, H - 225, f"HRA            :  Rs {d['hra']:>10,.2f}")
    c.drawString(40, H - 240, f"Allowances     :  Rs {d['allowance']:>10,.2f}")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(40, H - 260, f"Gross Salary   :  Rs {d['gross']:>10,.2f}")

    # Deductions
    c.setFont("Helvetica-Bold", 11)
    c.drawString(310, H - 185, "Deductions")
    c.line(310, H - 190, W - 40, H - 190)
    c.setFont("Helvetica", 10)
    c.drawString(310, H - 210, f"Provident Fund :  Rs {d['pf']:>10,.2f}")
    c.drawString(310, H - 225, f"ESI            :  Rs {d['esi']:>10,.2f}")
    c.drawString(310, H - 240, f"Prof. Tax (PT) :  Rs {d['pt']:>10,.2f}")
    c.drawString(310, H - 255, f"TDS            :  Rs {d['tds']:>10,.2f}")
    c.drawString(310, H - 270, f"LOP Deduction  :  Rs {d['lop']:>10,.2f}")

    # Net salary
    c.setFillColorRGB(0.07, 0.15, 0.40)
    c.rect(40, H - 315, W - 80, 30, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50,    H - 300, "NET SALARY")
    c.drawRightString(W - 50, H - 300, f"Rs {d['net']:,.2f}")

    c.setFillColorRGB(0.5, 0.5, 0.5)
    c.setFont("Helvetica", 8)
    c.drawCentredString(W / 2, 30,
        "Computer generated payslip — no signature required.")
    c.save()

    return send_file(file_path, as_attachment=True,
                     download_name=f"Payslip_{d['name']}_{d['month']}.pdf")


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
        SELECT e.name, p.month, p.gross, p.pf, p.esi, p.pt, p.tds, p.net
        FROM payroll p JOIN employees e ON e.id = p.emp_id
        ORDER BY p.month DESC
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
        SELECT e.name, e.bank_account, e.ifsc, p.net
        FROM payroll p JOIN employees e ON e.id = p.emp_id
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
        SELECT e.name, p.pf, p.esi
        FROM payroll p JOIN employees e ON e.id = p.emp_id
    """)
    data = cursor.fetchall()
    conn.close()
    return render_template("statutory_report.html", data=data)


# ──────────────────────────────────────────────────────────────
# Excel exports
# ──────────────────────────────────────────────────────────────
@app.route("/export_salary_register")
@role_required(["admin", "accountant"])
def export_salary_register():
    conn   = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT e.name, p.month, p.gross, p.pf, p.esi, p.pt, p.tds, p.net
        FROM payroll p JOIN employees e ON e.id = p.emp_id
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
        SELECT e.name, e.bank_account, e.ifsc, p.net
        FROM payroll p JOIN employees e ON e.id = p.emp_id
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
        """, (request.form["emp_id"],
              request.form["leave_date"],
              request.form["leave_type"]))
        conn.commit()
        conn.close()
        flash("Leave recorded.", "success")
        return redirect("/leave")

    cursor.execute("SELECT id, name FROM employees ORDER BY name")
    employees = cursor.fetchall()
    cursor.execute("""
        SELECT l.*, e.name FROM leaves l
        JOIN employees e ON e.id = l.emp_id
        ORDER BY l.leave_date DESC
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
        """, (request.form["emp_id"],
              request.form["ot_date"],
              float(request.form.get("hours", 0))))
        conn.commit()
        conn.close()
        flash("Overtime recorded.", "success")
        return redirect("/overtime")

    cursor.execute("SELECT id, name FROM employees ORDER BY name")
    employees = cursor.fetchall()
    cursor.execute("""
        SELECT o.*, e.name FROM overtime o
        JOIN employees e ON e.id = o.emp_id
        ORDER BY o.ot_date DESC
    """)
    records = cursor.fetchall()
    conn.close()
    return render_template("overtime.html",
                           employees=employees, records=records)


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
            INSERT INTO salary_history
                (emp_id, effective_date, basic, hra, allowances)
            VALUES (?,?,?,?,?)
        """, (request.form["emp_id"],
              request.form["date"],
              float(request.form.get("basic", 0)),
              float(request.form.get("hra", 0)),
              float(request.form.get("allowances", 0))))
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
    from database import DB_PATH
    timestamp   = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = os.path.join(BACKUP_DIR, f"payroll_{timestamp}.db")
    shutil.copy(DB_PATH, backup_file)
    flash(f"Backup created: payroll_{timestamp}.db", "success")
    return redirect("/backup_page")


@app.route("/restore/<filename>")
@role_required(["admin"])
def restore(filename):
    from database import DB_PATH
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
    files = sorted(os.listdir(BACKUP_DIR), reverse=True) \
            if os.path.exists(BACKUP_DIR) else []
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
    """, (request.form["name"], request.form.get("address", ""),
          request.form.get("pf", ""), request.form.get("esi", "")))
    conn.commit()
    conn.close()
    flash("Company added.", "success")
    return redirect("/companies")


# ──────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app.run(debug=True)
