import sqlite3
import os


def get_db_path():
    """
    On Render: persistent disk is mounted at /opt/render/project/src/data
    Locally  : use ./data next to this file
    We detect Render by the RENDER environment variable Render sets automatically.
    """
    if os.environ.get("RENDER"):
        base = "/opt/render/project/src/data"
    else:
        base = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")

    os.makedirs(base, exist_ok=True)
    return os.path.join(base, "payroll.db")


# Evaluated once at import time – safe for both gunicorn and `python app.py`
DB_PATH = get_db_path()


def init_db():
    """Create all tables and seed default data if the DB is brand-new."""
    print(f"[DB] Connecting to: {DB_PATH}")
    conn   = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.executescript("""
        PRAGMA journal_mode=WAL;

        CREATE TABLE IF NOT EXISTS users (
            id       INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            password TEXT NOT NULL,
            role     TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS companies (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            name       TEXT NOT NULL,
            address    TEXT,
            pf_number  TEXT,
            esi_number TEXT
        );

        CREATE TABLE IF NOT EXISTS employees (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id   INTEGER,
            emp_code     TEXT,
            name         TEXT NOT NULL,
            department   TEXT,
            designation  TEXT,
            doj          TEXT,
            basic        REAL DEFAULT 0,
            hra          REAL DEFAULT 0,
            allowance    REAL DEFAULT 0,
            bank_account TEXT,
            ifsc         TEXT,
            pan          TEXT,
            uan          TEXT,
            FOREIGN KEY (company_id) REFERENCES companies(id)
        );

        CREATE TABLE IF NOT EXISTS attendance (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            emp_id       INTEGER,
            month        TEXT,
            working_days INTEGER DEFAULT 26,
            present_days INTEGER DEFAULT 0,
            lop_days     INTEGER DEFAULT 0,
            FOREIGN KEY (emp_id) REFERENCES employees(id)
        );

        CREATE TABLE IF NOT EXISTS payroll (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            emp_id    INTEGER,
            month     TEXT,
            basic     REAL DEFAULT 0,
            hra       REAL DEFAULT 0,
            allowance REAL DEFAULT 0,
            gross     REAL DEFAULT 0,
            pf        REAL DEFAULT 0,
            esi       REAL DEFAULT 0,
            pt        REAL DEFAULT 0,
            tds       REAL DEFAULT 0,
            lop       REAL DEFAULT 0,
            net       REAL DEFAULT 0,
            FOREIGN KEY (emp_id) REFERENCES employees(id)
        );

        CREATE TABLE IF NOT EXISTS salary_history (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            emp_id         INTEGER,
            effective_date TEXT,
            basic          REAL DEFAULT 0,
            hra            REAL DEFAULT 0,
            allowances     REAL DEFAULT 0,
            FOREIGN KEY (emp_id) REFERENCES employees(id)
        );

        CREATE TABLE IF NOT EXISTS overtime (
            id      INTEGER PRIMARY KEY AUTOINCREMENT,
            emp_id  INTEGER,
            ot_date TEXT,
            hours   REAL DEFAULT 0,
            FOREIGN KEY (emp_id) REFERENCES employees(id)
        );

        CREATE TABLE IF NOT EXISTS leaves (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            emp_id     INTEGER,
            leave_date TEXT,
            leave_type TEXT,
            FOREIGN KEY (emp_id) REFERENCES employees(id)
        );
    """)

    # Seed default users only if none exist
    cursor.execute("SELECT COUNT(*) FROM users")
    if cursor.fetchone()[0] == 0:
        cursor.executemany(
            "INSERT INTO users (username, password, role) VALUES (?,?,?)",
            [
                ("admin",      "admin123", "admin"),
                ("hr",         "hr123",    "hr"),
                ("accountant", "acc123",   "accountant"),
            ]
        )
        print("[DB] Default users seeded.")

    # Seed default company only if none exist
    cursor.execute("SELECT COUNT(*) FROM companies")
    if cursor.fetchone()[0] == 0:
        cursor.execute("""
            INSERT INTO companies (name, address, pf_number, esi_number)
            VALUES ('My Company Pvt Ltd', '123 Business Park, Hyderabad',
                    'PF001', 'ESI001')
        """)
        print("[DB] Default company seeded.")

    conn.commit()
    conn.close()
    print(f"[DB] Ready ✅")
