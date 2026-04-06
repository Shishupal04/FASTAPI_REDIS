import sqlite3
import os


def _find_writable_db_path():
    """
    Try candidate paths in priority order.
    Returns the first path where we can actually write a SQLite file.
    This avoids all env-var guessing — we just probe at runtime.
    """
    candidates = [
        # 1. Render persistent disk (paid tier)
        "/opt/render/project/src/data/payroll.db",
        # 2. /tmp – always writable on every platform (Render free, Heroku, etc.)
        "/tmp/payroll_data/payroll.db",
        # 3. Local dev fallback
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "payroll.db"),
    ]

    for path in candidates:
        try:
            folder = os.path.dirname(path)
            os.makedirs(folder, exist_ok=True)
            # Actually try opening a real SQLite connection to confirm writable
            conn = sqlite3.connect(path)
            conn.execute("PRAGMA journal_mode=WAL")
            conn.close()
            print(f"[DB] Using path: {path}")
            return path
        except Exception as e:
            print(f"[DB] Skipping {path}: {e}")
            continue

    raise RuntimeError("No writable location found for the database!")


DB_PATH = _find_writable_db_path()


def get_db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row   # allows dict-style access
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def init_db():
    """Create all tables and seed initial data. Safe to call multiple times."""
    conn   = get_db()
    cursor = conn.cursor()

    cursor.executescript("""
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
            company_id   INTEGER DEFAULT 1,
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
            uan          TEXT
        );

        CREATE TABLE IF NOT EXISTS attendance (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            emp_id       INTEGER NOT NULL,
            month        TEXT NOT NULL,
            working_days INTEGER DEFAULT 26,
            present_days INTEGER DEFAULT 0,
            lop_days     INTEGER DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS payroll (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            emp_id    INTEGER NOT NULL,
            month     TEXT NOT NULL,
            basic     REAL DEFAULT 0,
            hra       REAL DEFAULT 0,
            allowance REAL DEFAULT 0,
            gross     REAL DEFAULT 0,
            pf        REAL DEFAULT 0,
            esi       REAL DEFAULT 0,
            pt        REAL DEFAULT 0,
            tds       REAL DEFAULT 0,
            lop       REAL DEFAULT 0,
            net       REAL DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS salary_history (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            emp_id         INTEGER NOT NULL,
            effective_date TEXT NOT NULL,
            basic          REAL DEFAULT 0,
            hra            REAL DEFAULT 0,
            allowances     REAL DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS overtime (
            id      INTEGER PRIMARY KEY AUTOINCREMENT,
            emp_id  INTEGER NOT NULL,
            ot_date TEXT NOT NULL,
            hours   REAL DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS leaves (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            emp_id     INTEGER NOT NULL,
            leave_date TEXT NOT NULL,
            leave_type TEXT
        );
    """)

    # Seed default users
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
        print("[DB] Default users created.")

    # Seed default company
    cursor.execute("SELECT COUNT(*) FROM companies")
    if cursor.fetchone()[0] == 0:
        cursor.execute("""
            INSERT INTO companies (name, address, pf_number, esi_number)
            VALUES ('My Company Pvt Ltd',
                    '123 Business Park, Hyderabad', 'PF001', 'ESI001')
        """)
        print("[DB] Default company created.")

    conn.commit()
    conn.close()
    print(f"[DB] Initialised successfully at {DB_PATH}")
