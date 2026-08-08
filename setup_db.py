import sqlite3
import hashlib

def hash_password(pw):
    return hashlib.sha256(pw.encode()).hexdigest()

def create_database():
    conn = sqlite3.connect('dashboard_data.db')
    c = conn.cursor()

    # Users table with role + view_only flag + allowed partners (comma-separated or 'All')
    c.execute('''CREATE TABLE IF NOT EXISTS Users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        role TEXT NOT NULL CHECK(role IN ('Admin','User')),
        allowed_partners TEXT NOT NULL DEFAULT 'All',
        view_only INTEGER NOT NULL DEFAULT 1
    )''')

    # Exchange rates table
    c.execute('''CREATE TABLE IF NOT EXISTS Exchange_Rates (
        currency_code TEXT PRIMARY KEY,
        rate_to_btn REAL NOT NULL,
        last_updated TEXT DEFAULT CURRENT_DATE
    )''')

    # Main flattened reports table
    c.execute('''CREATE TABLE IF NOT EXISTS Reports (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        instrument_id TEXT,
        title TEXT,
        creditor TEXT,
        agreement_structure TEXT,
        main_implementing_agency TEXT,
        measures TEXT,
        currency TEXT,
        disbursement_year TEXT,
        amount REAL,
        fiscal_year TEXT,
        upload_batch TEXT
    )''')

    # Audit log
    c.execute('''CREATE TABLE IF NOT EXISTS Audit_Log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT,
        action TEXT,
        detail TEXT,
        timestamp TEXT DEFAULT CURRENT_TIMESTAMP
    )''')

    # Uploaded file batches
    c.execute('''CREATE TABLE IF NOT EXISTS Upload_Batches (
        batch_id TEXT PRIMARY KEY,
        file_name TEXT,
        file_currency TEXT,
        uploaded_by TEXT,
        uploaded_at TEXT DEFAULT CURRENT_TIMESTAMP,
        row_count INTEGER DEFAULT 0,
        instrument_count INTEGER DEFAULT 0
    )''')

    # Default admin (password: admin123)
    c.execute('''INSERT OR IGNORE INTO Users (username, password, role, allowed_partners, view_only)
                 VALUES (?, ?, 'Admin', 'All', 0)''',
              ('admin', hash_password('admin123')))

    # Default exchange rates
    rates = [
        ('USD', 83.0), ('EUR', 90.5), ('INR', 1.0),
        ('JPY', 0.56), ('GBP', 105.0), ('CHF', 93.0),
        ('AUD', 54.0), ('SDR', 110.0)
    ]
    for code, rate in rates:
        c.execute('INSERT OR IGNORE INTO Exchange_Rates (currency_code, rate_to_btn) VALUES (?,?)',
                  (code, rate))

    conn.commit()
    conn.close()
    print("✅ Database created: dashboard_data.db")
    print("   Admin login → username: admin | password: admin123")

if __name__ == "__main__":
    create_database()
