import os
import io
import json
import streamlit as st
import pandas as pd
import numpy as np
import psycopg2
from sqlalchemy import create_engine
from urllib.parse import quote_plus
import plotly.express as px
import hashlib
import re
import traceback
from datetime import datetime, date, time as dt_time

# ─────────────────────────────────────────────
# OPENPYXL BUG FIX
# ─────────────────────────────────────────────
import openpyxl.descriptors.base
_original_convert = openpyxl.descriptors.base._convert

def _patched_convert(expected_type, value):
    if expected_type is datetime and isinstance(value, date) and not isinstance(value, datetime):
        value = datetime.combine(value, dt_time.min)
    return _original_convert(expected_type, value)

openpyxl.descriptors.base._convert = _patched_convert
 
# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="DCDMD Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

COLORS = {
    'loan':  '#E74C3C',
    'grant': '#2ECC71',
    'other': '#95A5A6',
    'card':  '#1E2130',
}


# ─────────────────────────────────────────────
# RAW FILE LOADER & FORMAT DETECTION
# ─────────────────────────────────────────────
def load_raw(file_obj):
    name = getattr(file_obj, 'name', '').lower()
    if name.endswith('.csv'):
        for enc in ('utf-8', 'cp1252', 'latin1'):
            try:
                if hasattr(file_obj, 'seek'): file_obj.seek(0)
                return pd.read_csv(file_obj, header=None, dtype=str, encoding=enc)
            except UnicodeDecodeError:
                continue
        raise UnicodeDecodeError('Unable to decode CSV with common encodings', b'', 0, 1, 'encoding')
    else:
        if hasattr(file_obj, 'seek'): file_obj.seek(0)
        return pd.read_excel(file_obj, sheet_name=0, header=None, dtype=str)


def detect_format(df_raw):
    for i in range(min(15, len(df_raw))):
        row_vals = [str(v).strip() for v in df_raw.iloc[i].values if pd.notna(v)]
        if 'Instrument Id' in row_vals:
            header_row = i
            headers = df_raw.iloc[i].tolist()
            col_map = {}
            for idx, h in enumerate(headers):
                h_s = str(h).strip()
                if h_s not in ('nan', 'NaN', 'None', ''):
                    col_map[h_s] = idx
            year_cols = {}
            for h, idx in col_map.items():
                m = re.fullmatch(r'(19[0-9]{2}|20[0-4][0-9]|2050)', h.strip())
                if m:
                    year_cols[int(m.group(0))] = idx
            fmt = 'flat' if i == 0 else 'staircase'
            return fmt, header_row, col_map, year_cols
    raise ValueError("Could not find 'Instrument Id' header in first 15 rows.")


def parse_flat(df_raw, header_row, col_map, year_cols, user_currency, capture_disbursements=True):
    data = df_raw.iloc[header_row + 1:].reset_index(drop=True)

    def gi(*names):
        for n in names:
            if n in col_map: return col_map[n]
        return None

    id_idx    = gi('Instrument Id')
    title_idx = gi('Title')
    main_title_idx = gi('Main Instrument Title')
    struct_idx = gi('Agreement Structure')
    cred_idx   = gi('Creditor Name', 'Creditor')
    adate_idx  = gi('Agreement Date')
    mdate_idx  = gi('Maturity Date')
    amt_idx    = gi('Amount')
    rev_idx    = gi('Revised Amount')
    agency_idx = gi('Main Implementing Agency Name', 'Main Implementing Agency')

    projects = {}
    disbursements = []

    for _, row in data.iterrows():
        raw_id = clean_val(row[id_idx]) if id_idx is not None else None
        if not raw_id: continue
        inst_id = raw_id.rstrip('.0') if raw_id.endswith('.0') else raw_id

        title = clean_val(row[title_idx]) if title_idx is not None else None
        if not title and main_title_idx is not None:
            title = clean_val(row[main_title_idx])

        creditor = None
        if cred_idx is not None:
            raw_c = clean_val(row[cred_idx])
            if raw_c:
                creditor = re.sub(r'\s*\(.*?\)\s*$', '', raw_c).strip()

        p = {
            'instrument_id': inst_id,
            'title': title,
            'agreement_structure': clean_val(row[struct_idx]) if struct_idx is not None else None,
            'creditor': creditor,
            'agreement_date': clean_date(row[adate_idx]) if adate_idx is not None else None,
            'maturity_date': clean_date(row[mdate_idx]) if mdate_idx is not None else None,
            'amount': clean_num(row[amt_idx]) if amt_idx is not None else None,
            'revised_amount': clean_num(row[rev_idx]) if rev_idx is not None else None,
            'economic_sector': None,
            'instrument_fund_use': None,
            'main_implementing_agency': clean_val(row[agency_idx]) if agency_idx is not None else None,
            'currency': user_currency,
        }
        projects[inst_id] = p

        if capture_disbursements and year_cols:
            for yr, yr_idx in year_cols.items():
                amt = clean_num(row[yr_idx])
                if amt and amt != 0:
                    disbursements.append({'instrument_id': inst_id, 'year': yr, 'amount': amt})

    df_p = pd.DataFrame(list(projects.values()))
    df_d = pd.DataFrame(disbursements)
    if not df_d.empty:
        df_d = df_d.groupby(['instrument_id', 'year'])['amount'].sum().reset_index()
    return df_p, df_d


def parse_upload_file(file_obj, user_currency, capture_disbursements=True):
    if hasattr(file_obj, 'seek'): file_obj.seek(0)
    df_raw = load_raw(file_obj)
    fmt, header_row, col_map, year_cols = detect_format(df_raw)
    if fmt == 'flat':
        df_p, df_d = parse_flat(df_raw, header_row, col_map, year_cols, user_currency, capture_disbursements)
    else:
        df_p, df_d, _ = parse_aggregate_wizard(file_obj, user_currency)
    return df_p, df_d, fmt, year_cols


# ─────────────────────────────────────────────
# DATABASE INIT — Postgres (Supabase), sqlite3-compatible wrapper
# ─────────────────────────────────────────────
@st.cache_resource(show_spinner=False)
def get_engine():
    """One pooled SQLAlchemy engine per running app process. Reused across
    reruns and sessions so we don't open a fresh socket to Supabase on every
    interaction."""
    cfg = st.secrets["postgres"]
    url = (
        f"postgresql+psycopg2://{cfg['user']}:{quote_plus(str(cfg['password']))}"
        f"@{cfg['host']}:{cfg['port']}/{cfg['dbname']}"
    )
    return create_engine(url, pool_pre_ping=True, pool_size=5, max_overflow=10)


class _CursorWrapper:
    """Makes a psycopg2 cursor accept the same `?`-style SQL the app already
    uses everywhere, by translating to psycopg2's `%s` placeholder style."""
    def __init__(self, cur):
        self._cur = cur

    def execute(self, sql, params=None):
        self._cur.execute(sql.replace('?', '%s'), params)
        return self

    def executemany(self, sql, seq_of_params):
        self._cur.executemany(sql.replace('?', '%s'), seq_of_params)
        return self

    def __getattr__(self, name):
        return getattr(self._cur, name)

    def __iter__(self):
        return iter(self._cur)


class PGConn:
    """Thin sqlite3.Connection-compatible wrapper around a psycopg2
    connection, so `conn.execute(...)`, `conn.executemany(...)`,
    `conn.cursor()`, and `with get_conn() as conn:` all keep working exactly
    as they did against SQLite."""
    def __init__(self, raw_conn):
        object.__setattr__(self, '_conn', raw_conn)

    def cursor(self):
        return _CursorWrapper(self._conn.cursor())

    def execute(self, sql, params=None):
        cur = self.cursor()
        cur.execute(sql, params)
        return cur

    def executemany(self, sql, seq_of_params):
        cur = self.cursor()
        cur.executemany(sql, seq_of_params)
        return cur

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        if exc_type is None:
            self._conn.commit()
        else:
            self._conn.rollback()
        return False

    def __getattr__(self, name):
        return getattr(self._conn, name)


def get_conn():
    """Pulls a connection from the pooled engine. Call .close() when done
    (or use `with get_conn() as conn:`) — this returns it to the pool rather
    than tearing down the socket."""
    return PGConn(get_engine().raw_connection())


def export_backup_json():
    """Snapshot of the core tables as JSON, used to give the admin a
    downloadable safety copy before a destructive delete/wipe. (Postgres on
    Supabase is durable on its own — this is a convenience export, not a
    substitute for Supabase's own backups.)"""
    engine = get_engine()
    payload = {}
    for tbl in ['Projects', 'Disbursements', 'Upload_Batches']:
        df = pd.read_sql(f"SELECT * FROM {tbl}", engine)
        payload[tbl] = df.to_dict(orient='records')
    return json.dumps(payload, default=str, indent=2).encode('utf-8')


def init_db():
    conn = get_conn()
    conn.execute('''CREATE TABLE IF NOT EXISTS Users (
        id SERIAL PRIMARY KEY,
        username TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        role TEXT NOT NULL,
        allowed_partners TEXT NOT NULL DEFAULT 'All',
        view_only INTEGER NOT NULL DEFAULT 1,
        hidden_columns TEXT DEFAULT '',
        hidden_projects TEXT DEFAULT '',
        hidden_partners TEXT DEFAULT ''
    )''')
    conn.execute('''
        INSERT INTO Users (username,password,role,allowed_partners,view_only)
        VALUES (?,?,?,?,?)
        ON CONFLICT (username) DO NOTHING
    ''', ('admin', hashlib.sha256('admin'.encode()).hexdigest(), 'Admin', 'All', 0))

    conn.execute('''CREATE TABLE IF NOT EXISTS Audit_Log (
        id SERIAL PRIMARY KEY,
        username TEXT, action TEXT, detail TEXT,
        timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )''')

    conn.execute('''CREATE TABLE IF NOT EXISTS FYP_Config (
        id SERIAL PRIMARY KEY,
        name TEXT UNIQUE NOT NULL,
        start_year INTEGER NOT NULL,
        end_year INTEGER NOT NULL
    )''')
    if conn.execute("SELECT COUNT(*) FROM FYP_Config").fetchone()[0] == 0:
        conn.executemany("INSERT INTO FYP_Config (name,start_year,end_year) VALUES (?,?,?)", [
            ("10th FYP (2008-2013)", 2008, 2012),
            ("11th FYP (2013-2018)", 2013, 2017),
            ("12th FYP (2018-2023)", 2018, 2022),
            ("Transition Period (2023-2024)", 2023, 2023),
            ("13th FYP (2024-2029)", 2024, 2029),
        ])

    conn.execute('''CREATE TABLE IF NOT EXISTS Upload_Batches (
        upload_id TEXT PRIMARY KEY,
        filename TEXT,
        upload_date TEXT,
        uploaded_by TEXT,
        currency TEXT,
        record_count INTEGER DEFAULT 0
    )''')

    conn.execute('''CREATE TABLE IF NOT EXISTS Exchange_Rates (
        currency_code TEXT PRIMARY KEY,
        rate_to_btn REAL NOT NULL,
        last_updated TEXT
    )''')
    if conn.execute("SELECT COUNT(*) FROM Exchange_Rates").fetchone()[0] == 0:
        conn.executemany('''
            INSERT INTO Exchange_Rates (currency_code,rate_to_btn,last_updated)
            VALUES (?,?,?)
            ON CONFLICT (currency_code) DO NOTHING
        ''', [
            ('BTN', 1.0, date.today().isoformat()),
            ('INR', 1.0, date.today().isoformat()),
            ('USD', 83.0, date.today().isoformat()),
            ('EUR', 90.5, date.today().isoformat()),
            ('XDR', 110.0, date.today().isoformat()),
        ])

    conn.execute('''CREATE TABLE IF NOT EXISTS Projects (
        instrument_id TEXT PRIMARY KEY,
        title TEXT,
        agreement_structure TEXT,
        creditor TEXT,
        agreement_date TEXT,
        maturity_date TEXT,
        amount REAL,
        revised_amount REAL,
        economic_sector TEXT,
        instrument_fund_use TEXT,
        main_implementing_agency TEXT,
        currency TEXT,
        upload_id TEXT
    )''')

    conn.execute('''CREATE TABLE IF NOT EXISTS Disbursements (
        instrument_id TEXT NOT NULL,
        year INTEGER NOT NULL,
        amount REAL NOT NULL DEFAULT 0,
        upload_id TEXT,
        PRIMARY KEY (instrument_id, year)
    )''')

    for alter_sql in [
        "ALTER TABLE Projects ADD COLUMN IF NOT EXISTS upload_id TEXT",
        "ALTER TABLE Projects ADD COLUMN IF NOT EXISTS file_type TEXT",
        "ALTER TABLE Disbursements ADD COLUMN IF NOT EXISTS upload_id TEXT",
        "ALTER TABLE Disbursements ADD COLUMN IF NOT EXISTS file_type TEXT",
        "ALTER TABLE Upload_Batches ADD COLUMN IF NOT EXISTS currency TEXT",
        "ALTER TABLE Upload_Batches ADD COLUMN IF NOT EXISTS record_count INTEGER DEFAULT 0",
        "ALTER TABLE Upload_Batches ADD COLUMN IF NOT EXISTS file_type TEXT",
        "ALTER TABLE Users ADD COLUMN IF NOT EXISTS hidden_partners TEXT DEFAULT ''",
    ]:
        try:
            conn.execute(alter_sql)
        except Exception:
            conn.rollback()

    try:
        conn.execute("CREATE INDEX IF NOT EXISTS idx_disburse_instrument_year ON Disbursements(instrument_id, year)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_projects_currency ON Projects(currency)")
    except Exception:
        conn.rollback()

    conn.commit()
    conn.close()

init_db()


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def hash_pw(pw):
    return hashlib.sha256(str(pw).encode()).hexdigest()


def log_action(username, action, detail=""):
    try:
        with get_conn() as conn:
            conn.execute("INSERT INTO Audit_Log (username,action,detail) VALUES (?,?,?)",
                         (username, action, detail))
    except Exception:
        pass


def clean_val(v):
    if v is None: return None
    s = str(v).strip()
    return None if s in ('', 'nan', 'NaN', 'None', '-', 'N/A', '#REF!') else s


def clean_num(v):
    s = clean_val(v)
    if s is None: return None
    try:
        return float(s.replace(',', ''))
    except (ValueError, AttributeError):
        return None


def clean_date(v):
    if v is None: return None
    if isinstance(v, (datetime, date)):
        return v.strftime('%Y-%m-%d')
    s = clean_val(v)
    if s is None: return None
    for fmt in ('%Y-%m-%d', '%d-%m-%Y', '%d/%m/%Y', '%Y/%m/%d', '%m/%d/%Y', '%d.%m.%Y'):
        try:
            return datetime.strptime(s, fmt).strftime('%Y-%m-%d')
        except ValueError:
            continue
    try:
        parsed = pd.to_datetime(s, errors='coerce')
        return parsed.strftime('%Y-%m-%d') if pd.notna(parsed) else None
    except Exception:
        return None


def classify(s):
    s = str(s or '').lower()
    if 'grant' in s: return 'Grant'
    if 'loan'  in s: return 'Loan'
    return 'Other'


def assign_fyp(yr, fyp_config):
    if pd.isna(yr): return "No Disbursements"
    try:
        y = int(yr)
        for _, r in fyp_config.iterrows():
            if r['start_year'] <= y <= r['end_year']:
                return r['name']
        return "Off-Cycle"
    except (ValueError, TypeError):
        return "Unknown"


# ─────────────────────────────────────────────
# PARSER — AGGREGATE WIZARD STAIRCASE FORMAT
# ─────────────────────────────────────────────
def parse_aggregate_wizard(file_obj, user_currency):
    if hasattr(file_obj, 'seek'): file_obj.seek(0)
    name = file_obj.name.lower()
    if name.endswith('.csv'):
        for enc in ('utf-8', 'cp1252', 'latin1'):
            try:
                if hasattr(file_obj, 'seek'): file_obj.seek(0)
                df_raw = pd.read_csv(file_obj, header=None, dtype=str, encoding=enc)
                break
            except UnicodeDecodeError:
                continue
    else:
        if hasattr(file_obj, 'seek'): file_obj.seek(0)
        df_raw = pd.read_excel(file_obj, sheet_name=0, header=None, dtype=str)

    header_row_idx = None
    for i in range(min(20, len(df_raw))):
        row_str = ' '.join([str(v) for v in df_raw.iloc[i].values])
        if 'Instrument Id' in row_str:
            header_row_idx = i
            break

    if header_row_idx is None:
        raise ValueError("Could not find 'Instrument Id' header in the first 20 rows.")

    headers = df_raw.iloc[header_row_idx].tolist()

    col_map = {}
    for idx, h in enumerate(headers):
        s = str(h).strip()
        if s and s not in ('nan', 'None'):
            col_map[s] = idx

    year_cols = {}
    for h, idx in col_map.items():
        m = re.fullmatch(r'(20[0-4][0-9]|2050)', h.strip())
        if m:
            year_cols[int(m.group(0))] = idx

    id_idx     = col_map.get('Instrument Id')
    title_idx  = col_map.get('Instrument Title') or col_map.get('Title') or col_map.get('Main Instrument Title')
    cred_idx   = col_map.get('Creditor Name') or col_map.get('Creditor')
    struct_idx = col_map.get('Agreement Structure')
    agency_idx = col_map.get('Main Implementing Agency Name') or col_map.get('Main Implementing Agency')
    meas_idx   = col_map.get('Measures')

    data = df_raw.iloc[header_row_idx + 1:].reset_index(drop=True)

    projects = {}
    disbursements = []
    current_id = None

    for row_i in range(len(data)):
        row = data.iloc[row_i]

        if id_idx is not None:
            raw_id = clean_val(row.iloc[id_idx])
            if raw_id:
                current_id = raw_id.rstrip('.0') if raw_id.endswith('.0') else raw_id
                if current_id not in projects:
                    projects[current_id] = {
                        'instrument_id': current_id,
                        'title': None,
                        'agreement_structure': None,
                        'creditor': None,
                        'agreement_date': None,
                        'maturity_date': None,
                        'amount': None,
                        'revised_amount': None,
                        'economic_sector': None,
                        'instrument_fund_use': None,
                        'main_implementing_agency': None,
                        'currency': user_currency,
                    }
                # Don't skip to the next row here — some exports (dense,
                # single-row-per-instrument layouts) put title/creditor/
                # agency/measures on this SAME row as the ID, not on
                # separate rows below it like the classic step-down
                # Aggregate Wizard layout does. Falling through lets both
                # layouts be read by the same field-filling logic below.

        if current_id is None:
            continue

        p = projects[current_id]

        if title_idx is not None and p['title'] is None:
            v = clean_val(row.iloc[title_idx])
            if v: p['title'] = v

        if cred_idx is not None and p['creditor'] is None:
            v = clean_val(row.iloc[cred_idx])
            if v:
                p['creditor'] = re.sub(r'\s*\(.*?\)\s*$', '', v).strip()

        if struct_idx is not None and p['agreement_structure'] is None:
            v = clean_val(row.iloc[struct_idx])
            if v: p['agreement_structure'] = v

        if agency_idx is not None and p['main_implementing_agency'] is None:
            v = clean_val(row.iloc[agency_idx])
            if v: p['main_implementing_agency'] = v

        if meas_idx is not None:
            measures_val = clean_val(row.iloc[meas_idx])
            if measures_val:
                for yr, yr_idx in year_cols.items():
                    amt = clean_num(row.iloc[yr_idx])
                    if amt is not None and amt != 0:
                        disbursements.append({
                            'instrument_id': current_id,
                            'year': yr,
                            'amount': amt
                        })

    df_projects = pd.DataFrame(list(projects.values()))
    df_disb = pd.DataFrame(disbursements)
    if not df_disb.empty:
        df_disb = df_disb.groupby(['instrument_id', 'year'])['amount'].sum().reset_index()

    return df_projects, df_disb, year_cols


# ─────────────────────────────────────────────
# CONFLICT DETECTION — compare new upload vs DB
# ─────────────────────────────────────────────
def check_conflicts(df_projects, df_disb):
    """
    Compare newly parsed file against what is already in the database.
    Returns (proj_conflicts_df, disb_conflicts_df).
    Each row describes one field/year where the new value differs from the stored one.
    """
    proj_conflicts = pd.DataFrame()
    disb_conflicts = pd.DataFrame()

    if df_projects.empty:
        return proj_conflicts, disb_conflicts

    ids = df_projects['instrument_id'].dropna().tolist()
    if not ids:
        return proj_conflicts, disb_conflicts

    try:
        conn = get_conn()
        cur = conn.cursor()

        # ── Project amount fields ────────────────────────────────────────────
        cur.execute(
            "SELECT instrument_id, amount, revised_amount FROM Projects "
            "WHERE instrument_id = ANY(%s)",
            (ids,)
        )
        existing_proj = pd.DataFrame(
            cur.fetchall(), columns=['instrument_id', 'amount', 'revised_amount']
        )

        # ── Disbursement amounts ─────────────────────────────────────────────
        cur.execute(
            "SELECT instrument_id, year, amount FROM Disbursements "
            "WHERE instrument_id = ANY(%s)",
            (ids,)
        )
        existing_disb = pd.DataFrame(
            cur.fetchall(), columns=['instrument_id', 'year', 'amount']
        )
        conn.close()

    except Exception:
        return proj_conflicts, disb_conflicts

    # ── Compare project amounts ──────────────────────────────────────────────
    if not existing_proj.empty:
        new_proj = df_projects[['instrument_id', 'amount', 'revised_amount']].copy()
        merged = new_proj.merge(existing_proj, on='instrument_id', suffixes=('_new', '_db'))

        rows = []
        for _, row in merged.iterrows():
            for field in ['amount', 'revised_amount']:
                new_v = row.get(f'{field}_new')
                db_v  = row.get(f'{field}_db')
                if (new_v is not None and db_v is not None
                        and pd.notna(new_v) and pd.notna(db_v)
                        and abs(float(new_v) - float(db_v)) > 0.01):
                    rows.append({
                        'Instrument ID': row['instrument_id'],
                        'Field': 'Original Amount' if field == 'amount' else 'Revised Amount',
                        'Stored Value': f"{db_v:,.2f}",
                        'New Value':    f"{new_v:,.2f}",
                    })
        if rows:
            proj_conflicts = pd.DataFrame(rows)

    # ── Compare disbursement amounts ─────────────────────────────────────────
    if not existing_disb.empty and not df_disb.empty:
        merged_d = df_disb.merge(
            existing_disb, on=['instrument_id', 'year'], suffixes=('_new', '_db')
        )
        rows_d = []
        for _, row in merged_d.iterrows():
            if abs(float(row['amount_new']) - float(row['amount_db'])) > 0.01:
                rows_d.append({
                    'Instrument ID':  row['instrument_id'],
                    'Year':           int(row['year']),
                    'Stored Amount':  f"{row['amount_db']:,.2f}",
                    'New Amount':     f"{row['amount_new']:,.2f}",
                })
        if rows_d:
            disb_conflicts = pd.DataFrame(rows_d)

    return proj_conflicts, disb_conflicts


# ─────────────────────────────────────────────
# SESSION STATE & LOGIN
# ─────────────────────────────────────────────
_defaults = {
    'logged_in': False, 'username': '', 'role': '',
    'allowed_partners': 'All', 'view_only': True,
    'hidden_cols': [], 'hidden_projs': [], 'hidden_partners': [],
}
for k, v in _defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


def show_login():
    c1, c2, c3 = st.columns([1, 1.5, 1])
    with c2:
        st.markdown("## 📊 DCDMD Dashboard")
        st.markdown("**Development Coordination & Debt Management Division**")
        st.markdown("---")
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        if st.button("Login", use_container_width=True, type="primary"):
            if not username or not password:
                return st.error("Enter both username and password.")
            conn = get_conn()
            row = conn.execute(
                "SELECT * FROM Users WHERE username=? AND password=?",
                (username, hash_pw(password))
            ).fetchone()
            conn.close()
            if row:
                st.session_state.logged_in        = True
                st.session_state.username         = row[1]
                st.session_state.role             = row[3]
                st.session_state.allowed_partners = row[4]
                st.session_state.view_only        = bool(row[5])
                st.session_state.hidden_cols      = [c for c in row[6].split(',') if c] if row[6] else []
                st.session_state.hidden_projs     = [p for p in row[7].split(',') if p] if row[7] else []
                st.session_state.hidden_partners  = [p for p in row[8].split(',') if p] if (len(row) > 8 and row[8]) else []
                log_action(username, "LOGIN")
                st.rerun()
            else:
                st.error("❌ Invalid credentials.")


def render_sidebar():
    st.sidebar.markdown(f"### 👤 {st.session_state.username}")
    st.sidebar.caption(
        f"Role: **{st.session_state.role}**  |  "
        f"{'🔒 View Only' if st.session_state.view_only else '✏️ Edit Access'}"
    )
    st.sidebar.markdown("---")

    if st.session_state.role == 'Admin':
        pages = ["📊 Dashboard", "💱 Exchange Rates", "📁 Upload & Manage Data",
                 "👥 User Management", "⚙️ System Settings", "📋 Audit Log"]
    else:
        pages = ["📊 Dashboard"]

    choice = st.sidebar.radio("Navigation", pages, label_visibility="collapsed")
    st.sidebar.markdown("---")
    if st.sidebar.button("🚪 Logout"):
        log_action(st.session_state.username, "LOGOUT")
        for k, v in _defaults.items(): st.session_state[k] = v
        st.rerun()
    return choice.split(" ", 1)[-1].strip()


# ─────────────────────────────────────────────
# PAGE: DASHBOARD
# ─────────────────────────────────────────────
def page_dashboard():
    st.title("📊 DCDMD Project Dashboard")

    engine = get_engine()
    df_all = pd.read_sql('''
        SELECT p.instrument_id, p.title, p.agreement_structure, p.creditor,
               p.agreement_date, p.maturity_date, p.amount, p.revised_amount,
               p.economic_sector, p.instrument_fund_use, p.main_implementing_agency,
               p.currency,
               d.year AS disbursement_year, d.amount AS disbursed_amount
        FROM Projects p
        LEFT JOIN Disbursements d ON p.instrument_id = d.instrument_id
    ''', engine)
    rates_df   = pd.read_sql("SELECT * FROM Exchange_Rates", engine).set_index('currency_code')
    fyp_config = pd.read_sql("SELECT * FROM FYP_Config ORDER BY start_year", engine)

    if df_all.empty:
        return st.warning("⚠️ No data yet. Admin must upload reports first.")

    # ── Security filters ─────────────────────────────────────────────────────
    if st.session_state.role != 'Admin' and st.session_state.hidden_projs:
        df_all = df_all[~df_all['instrument_id'].isin(st.session_state.hidden_projs)]

    if st.session_state.allowed_partners != 'All':
        allowed = [p.strip() for p in st.session_state.allowed_partners.split(',')]
        df_all  = df_all[df_all['creditor'].isin(allowed)]

    if st.session_state.role != 'Admin':
        hidden_partners = st.session_state.get('hidden_partners') or []
        if hidden_partners:
            hidden_set = {str(hp).strip().lower() for hp in hidden_partners if hp}
            def _is_hidden_partner(creditor_val):
                if creditor_val is None or (isinstance(creditor_val, float) and pd.isna(creditor_val)):
                    return False
                return str(creditor_val).strip().lower() in hidden_set
            mask = df_all['creditor'].apply(_is_hidden_partner)
            df_all = df_all[~mask]

    if df_all.empty:
        return st.warning("No data available for your account. Contact your administrator.")

    # ── Sidebar filters ──────────────────────────────────────────────────────
    st.sidebar.markdown("### 🔍 Filters")
    creditors  = ['All'] + sorted(df_all['creditor'].dropna().unique().tolist())
    agencies   = ['All'] + sorted(df_all['main_implementing_agency'].dropna().unique().tolist())
    structures = ['All'] + sorted(df_all['agreement_structure'].dropna().unique().tolist())

    sel_cred   = st.sidebar.selectbox("Development Partner", creditors)
    sel_agency = st.sidebar.selectbox("Implementing Agency", agencies)
    sel_struct = st.sidebar.selectbox("Instrument Type", structures)

    fyp_options = ["All FYPs"] + fyp_config['name'].tolist()
    sel_fyp = st.sidebar.selectbox("📂 Five Year Plan", fyp_options)

    all_years = sorted(df_all['disbursement_year'].dropna().unique().astype(int).tolist())
    fy_labels = {yr: f"FY {yr}/{str(yr + 1)[-2:]}" for yr in all_years}
    fy_options = ["All Fiscal Years"] + [fy_labels[yr] for yr in all_years]
    sel_fy = st.sidebar.selectbox("📅 Fiscal Year", fy_options)

    # ── Currency mode ─────────────────────────────────────────────────────────
    st.sidebar.markdown("### 💱 Currency View")
    currency_mode = st.sidebar.selectbox("Display In", [
        "Aggregate (BTN / Nu.)",
        "Instrument Currency (ForEx)",
    ])

    available_currencies = sorted(df_all['currency'].dropna().unique().tolist())
    sel_currency = None
    if currency_mode == "Instrument Currency (ForEx)":
        if available_currencies:
            sel_currency = st.sidebar.selectbox("Select Currency", available_currencies)
        else:
            st.sidebar.info("No currency info found.")

    # ── Apply filters ─────────────────────────────────────────────────────────
    dff = df_all.copy()
    if sel_cred   != 'All': dff = dff[dff['creditor'] == sel_cred]
    if sel_agency != 'All': dff = dff[dff['main_implementing_agency'] == sel_agency]
    if sel_struct != 'All': dff = dff[dff['agreement_structure'] == sel_struct]

    if sel_fyp != "All FYPs":
        fyp_row   = fyp_config[fyp_config['name'] == sel_fyp].iloc[0]
        fyp_start = int(fyp_row['start_year'])
        fyp_end   = int(fyp_row['end_year'])
        dff = dff[
            dff['disbursement_year'].isna() |
            dff['disbursement_year'].between(fyp_start, fyp_end)
        ]

    if sel_fy != "All Fiscal Years":
        sel_cal_yr = next(yr for yr, lbl in fy_labels.items() if lbl == sel_fy)
        dff = dff[
            dff['disbursement_year'].isna() |
            (dff['disbursement_year'] == sel_cal_yr)
        ]

    # ── Currency conversion ───────────────────────────────────────────────────
    if currency_mode == "Aggregate (BTN / Nu.)":
        missing = [c for c in dff['currency'].dropna().unique()
                   if str(c).upper() not in rates_df.index]
        if missing:
            st.warning(f"⚠️ Missing exchange rates for: **{', '.join(missing)}** — shown as 1:1. "
                       "Go to Exchange Rates page to add them.")

        def to_btn(row):
            code = str(row.get('currency', 'BTN') or 'BTN').strip().upper()
            amt  = float(row['disbursed_amount']) if pd.notna(row['disbursed_amount']) else 0.0
            if code == 'BTN': return amt
            rate = rates_df.loc[code, 'rate_to_btn'] if code in rates_df.index else 1.0
            return amt * rate

        dff['display_amount'] = dff.apply(to_btn, axis=1)
        unit = "Nu. (Millions)"
    else:
        if sel_currency:
            dff = dff[dff['currency'].astype(str).str.upper() == sel_currency.upper()]
            unit = f"{sel_currency} (Millions)"
        else:
            unit = "Millions"
        dff['display_amount'] = dff['disbursed_amount'].fillna(0)

    dff['type'] = dff['agreement_structure'].apply(classify)
    dff['fyp']  = dff['disbursement_year'].apply(lambda y: assign_fyp(y, fyp_config))

    # ── KPI Cards ─────────────────────────────────────────────────────────────
    total  = dff['display_amount'].sum()
    grants = dff[dff['type'] == 'Grant']['display_amount'].sum()
    loans  = dff[dff['type'] == 'Loan']['display_amount'].sum()
    n_proj = dff['instrument_id'].nunique()
    n_part = 1 if sel_cred != 'All' else dff['creditor'].nunique()

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("💰 Total Amount", f"{total:,.2f}",  unit)
    k2.metric("🤝 Total Grants", f"{grants:,.2f}", unit)
    k3.metric("🏦 Total Loans",  f"{loans:,.2f}",  unit)
    k4.metric("📁 Projects",     f"{n_proj}")
    k5.metric("🌐 Dev Partners", f"{n_part}")
    st.markdown("---")

    color_map = {
        'Grant': COLORS['grant'],
        'Loan':  COLORS['loan'],
        'Other': COLORS['other'],
    }

    # ── Row 1: FYP Bar + Yearly Timeline ─────────────────────────────────────
    col1, col2 = st.columns([1, 1.2])

    with col1:
        st.subheader("📋 Amount by Five-Year Plan")
        fyp_data = (dff[dff['display_amount'] > 0]
                    .groupby(['fyp', 'type'])['display_amount'].sum()
                    .reset_index())
        if not fyp_data.empty:
            fig = px.bar(fyp_data, x='fyp', y='display_amount', color='type',
                         barmode='group', text_auto='.2s',
                         color_discrete_map=color_map,
                         labels={'fyp': 'Plan Period', 'display_amount': unit, 'type': 'Type'})
            fig.update_layout(
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
                margin=dict(l=0, r=0, t=30, b=0), xaxis_tickangle=-30
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No disbursement data for selected filters.")

    with col2:
        st.subheader("📅 Yearly Amount Timeline")
        fyp_opts  = ["Show All"] + sorted([
            f for f in dff['fyp'].unique()
            if "No Disbursements" not in str(f) and "Unknown" not in str(f)
        ])
        link_sel  = st.selectbox("Zoom into FYP", fyp_opts, key="fyp_zoom")
        yearly_df = dff[dff['display_amount'] > 0].copy()
        if link_sel != "Show All":
            yearly_df = yearly_df[yearly_df['fyp'] == link_sel]
        yearly_data = (yearly_df.groupby(['disbursement_year', 'type'])['display_amount']
                       .sum().reset_index().sort_values('disbursement_year'))
        if not yearly_data.empty:
            fig2 = px.bar(yearly_data, x='disbursement_year', y='display_amount', color='type',
                          barmode='stack', color_discrete_map=color_map,
                          labels={'disbursement_year': 'Year', 'display_amount': unit, 'type': 'Type'})
            fig2.update_layout(
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
                margin=dict(l=0, r=0, t=10, b=0),
                xaxis=dict(tickmode='linear', tickformat='d', tickangle=-45)
            )
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("No disbursements for selected period.")

    st.markdown("---")

    # ── Row 2: Creditor Donut + Agency Bar ───────────────────────────────────
    col3, col4 = st.columns([1, 1.2])

    with col3:
        st.subheader("🌐 Creditor Share")
        cred_pie = (dff[dff['display_amount'] > 0]
                    .groupby('creditor')['display_amount'].sum().reset_index())
        if not cred_pie.empty:
            fig3 = px.pie(cred_pie, values='display_amount', names='creditor',
                          hole=0.42,
                          color_discrete_sequence=px.colors.qualitative.Pastel)
            fig3.update_traces(textposition='inside', textinfo='percent+label')
            fig3.update_layout(showlegend=False, margin=dict(l=0, r=0, t=10, b=0),
                               paper_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig3, use_container_width=True)

    with col4:
        st.subheader("🏭 Agency Allocation")
        agency_bar = (dff[dff['display_amount'] > 0]
                      .groupby(['main_implementing_agency', 'type'])['display_amount']
                      .sum().reset_index()
                      .sort_values('display_amount', ascending=True))
        if not agency_bar.empty:
            fig4 = px.bar(agency_bar, x='display_amount', y='main_implementing_agency',
                          color='type', barmode='stack', orientation='h',
                          color_discrete_map=color_map,
                          labels={'display_amount': unit, 'main_implementing_agency': 'Agency', 'type': 'Type'})
            fig4.update_layout(
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
                margin=dict(l=0, r=0, t=10, b=0), yaxis_title=""
            )
            st.plotly_chart(fig4, use_container_width=True)

    st.markdown("---")

    # ── Fund Use ──────────────────────────────────────────────────────────────
    st.subheader("💼 Fund Use Allocation")
    fund_df = (dff[dff['display_amount'] > 0]
               .dropna(subset=['instrument_fund_use'])
               .groupby('instrument_fund_use')['display_amount']
               .sum().reset_index())
    if not fund_df.empty:
        fig5 = px.pie(fund_df, values='display_amount', names='instrument_fund_use',
                      hole=0.42, color_discrete_sequence=px.colors.qualitative.Pastel)
        fig5.update_traces(textposition='inside', textinfo='percent+label')
        fig5.update_layout(showlegend=False, margin=dict(l=0, r=0, t=10, b=0),
                           paper_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig5, use_container_width=True)
    else:
        st.info("No fund use data for selected filters.")

    # ── Project Drill-Down ────────────────────────────────────────────────────
    st.subheader("🔍 Project Drill-Down")
    uniq = dff.drop_duplicates('instrument_id').copy()
    uniq['label'] = uniq['instrument_id'] + "  —  " + uniq['title'].fillna('Unknown')
    sel_label = st.selectbox("Search Project", ["— Select a project —"] + sorted(uniq['label'].tolist()))

    if sel_label != "— Select a project —":
        sel_id     = sel_label.split("  —  ")[0].strip()
        proj_rows  = dff[dff['instrument_id'] == sel_id]
        p          = proj_rows.iloc[0]
        total_disb = proj_rows['display_amount'].sum()
        ptype      = p['type']
        tc = COLORS['grant'] if ptype == 'Grant' else (COLORS['loan'] if ptype == 'Loan' else COLORS['other'])

        def safe_val(col_key):
            if st.session_state.role != 'Admin' and col_key in st.session_state.hidden_cols:
                return "🔒 Hidden"
            return p[col_key] if pd.notna(p.get(col_key)) else 'N/A'

        ci1, ci2 = st.columns([1, 1])
        with ci1:
            st.markdown(f"""
            <div style="background:{COLORS['card']};border-radius:12px;padding:20px;border-left:4px solid {tc}">
                <h4 style="margin:0 0 12px;color:#fff">{p.get('title','Unknown')}</h4>
                <table style="width:100%;color:#ccc;font-size:13px;border-collapse:collapse">
                    <tr><td><b>🆔 Instrument ID</b></td><td style="text-align:right">{p['instrument_id']}</td></tr>
                    <tr><td><b>🏷️ Type</b></td><td style="text-align:right"><span style="color:{tc};font-weight:bold">{ptype}</span></td></tr>
                    <tr><td><b>🤝 Creditor</b></td><td style="text-align:right">{p.get('creditor','N/A')}</td></tr>
                    <tr><td><b>🏛️ Agency</b></td><td style="text-align:right">{p.get('main_implementing_agency','N/A')}</td></tr>
                    <tr><td><b>📅 Agreement Date</b></td><td style="text-align:right">{safe_val('agreement_date')}</td></tr>
                    <tr><td><b>🏁 Maturity Date</b></td><td style="text-align:right">{safe_val('maturity_date')}</td></tr>
                    <tr><td><b>💳 Original Amount</b></td><td style="text-align:right">{safe_val('amount')}</td></tr>
                    <tr><td><b>🔄 Revised Amount</b></td><td style="text-align:right">{safe_val('revised_amount')}</td></tr>
                    <tr style="border-top:1px solid #444">
                        <td><b>💰 Total Disbursed</b></td>
                        <td style="text-align:right"><b style="color:#fff">{total_disb:,.2f} {unit}</b></td>
                    </tr>
                </table>
            </div>""", unsafe_allow_html=True)

        with ci2:
            yr_data = proj_rows.dropna(subset=['disbursement_year'])
            if not yr_data.empty:
                fig6 = px.bar(yr_data.sort_values('disbursement_year'),
                              x='disbursement_year', y='display_amount',
                              title="Year-wise Disbursements",
                              labels={'disbursement_year': 'Year', 'display_amount': unit},
                              color_discrete_sequence=[tc])
                fig6.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                                   margin=dict(l=0, r=0, t=40, b=0),
                                   xaxis=dict(tickmode='linear', tickformat='d'))
                st.plotly_chart(fig6, use_container_width=True)

    st.markdown("---")

    # ── Data Matrix Table ──────────────────────────────────────────────────────
    with st.expander("📄 Full Data Matrix", expanded=False):
        search = st.text_input("🔍 Search (ID, Title, Creditor, Agency):", "")

        base_cols = ['instrument_id', 'title', 'creditor', 'agreement_structure',
                     'main_implementing_agency', 'currency']
        secure_cols = ['agreement_date', 'maturity_date', 'amount', 'revised_amount']
        for sc in secure_cols:
            if st.session_state.role == 'Admin' or sc not in st.session_state.hidden_cols:
                base_cols.append(sc)

        df_base = dff[base_cols].drop_duplicates('instrument_id')

        if not dff['disbursement_year'].isna().all():
            df_years = (dff.dropna(subset=['disbursement_year'])
                        .pivot_table(index='instrument_id', columns='disbursement_year',
                                     values='display_amount', aggfunc='sum')
                        .reset_index())
            df_display = pd.merge(df_base, df_years, on='instrument_id', how='left')
        else:
            df_display = df_base.copy()

        rename_dict   = {}
        year_col_list = []
        for c in df_display.columns:
            cs = str(c)
            if re.fullmatch(r'\d{4}(\.0)?', cs.strip()):
                yr = str(int(float(cs)))
                rename_dict[c] = yr
                year_col_list.append(yr)
        df_display    = df_display.rename(columns=rename_dict)
        year_col_list = sorted(set(year_col_list))

        if year_col_list:
            df_display[year_col_list] = df_display[year_col_list].fillna(0)
            df_display['Total Disbursed'] = df_display[year_col_list].sum(axis=1)
            final_cols = base_cols + year_col_list + ['Total Disbursed']
        else:
            final_cols = base_cols

        final_cols = [c for c in final_cols if c in df_display.columns]
        df_display = df_display[final_cols]

        if search:
            mask = df_display.astype(str).apply(
                lambda x: x.str.contains(search, case=False, na=False)
            ).any(axis=1)
            df_display = df_display[mask]

        # ── Number formatting: comma-separated for all amount columns ─────────
        def _fmt_commas(v, decimals=0):
            if pd.isna(v):
                return ""
            try:
                return f"{float(v):,.{decimals}f}"
            except (ValueError, TypeError):
                return v

        df_display_fmt = df_display.copy()
        for yr_col in year_col_list:
            if yr_col in df_display_fmt.columns:
                df_display_fmt[yr_col] = df_display_fmt[yr_col].apply(lambda v: _fmt_commas(v, 0))
        for amt_col in ['amount', 'revised_amount', 'Total Disbursed']:
            if amt_col in df_display_fmt.columns:
                df_display_fmt[amt_col] = df_display_fmt[amt_col].apply(lambda v: _fmt_commas(v, 2))

        rename_display = {'amount': 'Original Amount', 'revised_amount': 'Revised Amount'}
        df_display_fmt = df_display_fmt.rename(columns=rename_display)

        st.dataframe(
            df_display_fmt,
            use_container_width=True,
            hide_index=True,
        )
        st.caption(f"Showing {len(df_display)} projects | Unit: {unit}")


# ─────────────────────────────────────────────
# PAGE: EXCHANGE RATES
# ─────────────────────────────────────────────
def page_rates():
    st.title("💱 Exchange Rates")
    st.caption("All amounts are converted to BTN using these rates. 1 BTN = 1 BTN and 1 INR = 1 BTN by default.")

    engine = get_engine()
    rates = pd.read_sql(
        "SELECT currency_code, rate_to_btn, last_updated FROM Exchange_Rates ORDER BY currency_code",
        engine
    )

    st.dataframe(
        rates.rename(columns={
            'currency_code': 'Currency Code',
            'rate_to_btn': '1 Unit = X BTN',
            'last_updated': 'Last Updated'
        }),
        use_container_width=True, hide_index=True
    )

    st.markdown("---")
    st.subheader("Add / Edit Rate")

    existing_codes = pd.read_sql("SELECT currency_code FROM Exchange_Rates", engine)['currency_code'].tolist()

    c1, c2, c3, c4 = st.columns([1, 1, 1, 1])
    with c1:
        edit_mode = st.selectbox("Mode", ["Add New", "Edit Existing"])
    with c2:
        if edit_mode == "Edit Existing":
            code = st.selectbox("Currency Code", existing_codes)
        else:
            code = st.text_input("Currency Code", placeholder="e.g. JPY").strip().upper()
    with c3:
        new_rate = st.number_input("Rate (1 unit = ? BTN)", min_value=0.0001, value=1.0, step=0.01, format="%.4f")
    with c4:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("💾 Save Rate", type="primary") and code:
            conn3 = get_conn()
            conn3.execute('''
                INSERT INTO Exchange_Rates (currency_code, rate_to_btn, last_updated) VALUES (?,?,?)
                ON CONFLICT(currency_code) DO UPDATE SET
                    rate_to_btn=excluded.rate_to_btn,
                    last_updated=excluded.last_updated
            ''', (code, new_rate, date.today().isoformat()))
            conn3.commit(); conn3.close()
            log_action(st.session_state.username, "RATE_UPDATE", f"1 {code} = {new_rate} BTN")
            st.success(f"✅ Saved: 1 {code} = {new_rate} BTN")
            st.rerun()

    if st.session_state.role == 'Admin' and existing_codes:
        st.markdown("---")
        st.subheader("Delete Rate")
        del_code = st.selectbox("Select currency to delete", existing_codes, key="del_rate_sel")
        if st.button("🗑️ Delete", type="secondary"):
            conn4 = get_conn()
            conn4.execute("DELETE FROM Exchange_Rates WHERE currency_code=?", (del_code,))
            conn4.commit(); conn4.close()
            log_action(st.session_state.username, "RATE_DELETE", del_code)
            st.success(f"Deleted {del_code}")
            st.rerun()


# ─────────────────────────────────────────────
# PAGE: UPLOAD & MANAGE DATA (Admin only)
# ─────────────────────────────────────────────
def page_upload():
    st.title("📁 Upload & Manage Data")
    st.info(
        "Upload Excel or CSV reports (Flat or Staircase). The system extracts instruments "
        "and disbursement data. Existing records are updated only where fields are blank — "
        "unless you choose to overwrite conflicting amounts."
    )

    st.markdown("### 1. Declare Currency & Upload File")
    st.caption("Select the currency that applies to ALL amounts in this file.")

    CURRENCY_LIST = ["INR", "USD", "BTN", "EUR", "XDR", "JPY", "KRW", "CAD", "DKK", "CHF", "GBP"]
    user_currency = st.selectbox("File Currency", CURRENCY_LIST)

    uploaded = st.file_uploader("Choose Excel or CSV file", type=['xlsx', 'xls', 'csv'])

    if uploaded:
        st.markdown("---")
        st.markdown("### 2. Preview & Confirm")

        try:
            df_projects, df_disb, detected_fmt, year_cols = parse_upload_file(
                uploaded, user_currency, capture_disbursements=True
            )

            st.success(
                f"✅ Parsed **{len(df_projects)} instruments** "
                f"and **{len(df_disb)} disbursement records**"
            )

            yr_list = sorted(year_cols.keys()) if year_cols else []
            if yr_list:
                st.info(
                    f"📅 Year columns detected: **{yr_list[0]} → {yr_list[-1]}** "
                    f"({len(yr_list)} years)"
                )

            with st.expander("👁️ Preview parsed instruments (first 10)"):
                st.dataframe(df_projects.head(10), use_container_width=True, hide_index=True)

            with st.expander("👁️ Preview disbursements (first 20)"):
                st.dataframe(
                    df_disb.head(20) if not df_disb.empty
                    else pd.DataFrame({"info": ["No disbursements found"]}),
                    use_container_width=True, hide_index=True
                )

            mismatch = False
            if 'currency' in df_projects.columns:
                file_currencies = sorted(set([
                    c for c in df_projects['currency'].dropna().unique()
                ]))
                if file_currencies and any(
                    str(c).upper() != user_currency.upper() for c in file_currencies
                ):
                    st.warning(
                        f"⚠️ Detected currencies in file: {file_currencies}. "
                        f"You chose {user_currency}."
                    )
                    mismatch = True

            # ── CONFLICT DETECTION ────────────────────────────────────────────
            st.markdown("---")
            st.markdown("### 3. Review Conflicts (if any)")

            proj_conflicts, disb_conflicts = check_conflicts(df_projects, df_disb)
            total_conflicts = len(proj_conflicts) + len(disb_conflicts)

            overwrite_amounts = False

            if total_conflicts == 0:
                st.success("✅ No amount conflicts found — all values in this file are either new or blank in the database.")
            else:
                st.warning(
                    f"⚠️ **{total_conflicts} amount conflict(s)** detected. "
                    "The values below already exist in the database but differ from this file."
                )

                if not proj_conflicts.empty:
                    st.markdown("**📋 Project Amount Conflicts:**")
                    st.dataframe(proj_conflicts, use_container_width=True, hide_index=True)

                if not disb_conflicts.empty:
                    st.markdown("**📅 Disbursement Amount Conflicts:**")
                    st.dataframe(disb_conflicts, use_container_width=True, hide_index=True)

                st.markdown("**What would you like to do with these conflicts?**")

                overwrite_choice = st.radio(
                    "Choose action for conflicting amounts:",
                    options=[
                        "🔒 Keep existing — do not overwrite stored amounts",
                        "🔄 Overwrite — replace stored amounts with new values from this file",
                    ],
                    index=0,
                    key="conflict_overwrite_choice",
                )
                overwrite_amounts = overwrite_choice.startswith("🔄")

                if overwrite_amounts:
                    st.warning(
                        "⚡ **Overwrite mode selected.** "
                        f"The {total_conflicts} conflicting amount(s) shown above "
                        "will be replaced with the new values from this file."
                    )
                else:
                    st.info(
                        "🔒 **Keep mode selected.** "
                        "Existing amounts will not be changed. "
                        "Only blank/null fields will be filled from this file."
                    )

            # ── SAVE ──────────────────────────────────────────────────────────
            st.markdown("---")
            st.markdown("### 4. Save to Database")

            confirm_continue = True
            if mismatch:
                confirm_continue = st.checkbox(
                    "I acknowledge the currency mismatch and want to continue",
                    value=False
                )

            if st.button("🔄 Smart Merge & Save", type="primary", use_container_width=True) and confirm_continue:
                with st.spinner("Merging data into database…"):
                    try:
                        batch_id = datetime.now().strftime("BATCH_%Y%m%d_%H%M%S")

                        proj_tuples = []
                        for _, r in df_projects.iterrows():
                            proj_tuples.append((
                                r.get('instrument_id'), r.get('title'),
                                r.get('agreement_structure'), r.get('creditor'),
                                r.get('agreement_date'), r.get('maturity_date'),
                                r.get('amount'), r.get('revised_amount'),
                                r.get('economic_sector'), r.get('instrument_fund_use'),
                                r.get('main_implementing_agency'),
                                r.get('currency') or user_currency,
                                batch_id, detected_fmt
                            ))

                        disb_tuples = []
                        if not df_disb.empty:
                            for _, r in df_disb.iterrows():
                                disb_tuples.append((
                                    r['instrument_id'], int(r['year']),
                                    float(r['amount']), batch_id, detected_fmt
                                ))

                        # ── SQL: NULL-fill only (default / keep mode) ─────────
                        project_sql_keep = '''
                            INSERT INTO Projects (
                                instrument_id, title, agreement_structure, creditor,
                                agreement_date, maturity_date, amount, revised_amount,
                                economic_sector, instrument_fund_use,
                                main_implementing_agency, currency, upload_id, file_type
                            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                            ON CONFLICT(instrument_id) DO UPDATE SET
                                title = CASE WHEN Projects.title IS NULL OR Projects.title='' THEN NULLIF(excluded.title,'') ELSE Projects.title END,
                                agreement_structure = CASE WHEN Projects.agreement_structure IS NULL OR Projects.agreement_structure='' THEN NULLIF(excluded.agreement_structure,'') ELSE Projects.agreement_structure END,
                                creditor = CASE WHEN Projects.creditor IS NULL OR Projects.creditor='' THEN NULLIF(excluded.creditor,'') ELSE Projects.creditor END,
                                agreement_date = CASE WHEN Projects.agreement_date IS NULL OR Projects.agreement_date='' THEN NULLIF(excluded.agreement_date,'') ELSE Projects.agreement_date END,
                                maturity_date = CASE WHEN Projects.maturity_date IS NULL OR Projects.maturity_date='' THEN NULLIF(excluded.maturity_date,'') ELSE Projects.maturity_date END,
                                amount = CASE WHEN Projects.amount IS NULL THEN excluded.amount ELSE Projects.amount END,
                                revised_amount = CASE WHEN Projects.revised_amount IS NULL THEN excluded.revised_amount ELSE Projects.revised_amount END,
                                economic_sector = CASE WHEN Projects.economic_sector IS NULL OR Projects.economic_sector='' THEN NULLIF(excluded.economic_sector,'') ELSE Projects.economic_sector END,
                                instrument_fund_use = CASE WHEN Projects.instrument_fund_use IS NULL OR Projects.instrument_fund_use='' THEN NULLIF(excluded.instrument_fund_use,'') ELSE Projects.instrument_fund_use END,
                                main_implementing_agency = CASE WHEN Projects.main_implementing_agency IS NULL OR Projects.main_implementing_agency='' THEN NULLIF(excluded.main_implementing_agency,'') ELSE Projects.main_implementing_agency END,
                                currency = CASE WHEN Projects.currency IS NULL OR Projects.currency='' THEN NULLIF(excluded.currency,'') ELSE Projects.currency END,
                                upload_id = excluded.upload_id,
                                file_type = excluded.file_type
                        '''

                        # ── SQL: overwrite amounts (user confirmed) ───────────
                        project_sql_overwrite = '''
                            INSERT INTO Projects (
                                instrument_id, title, agreement_structure, creditor,
                                agreement_date, maturity_date, amount, revised_amount,
                                economic_sector, instrument_fund_use,
                                main_implementing_agency, currency, upload_id, file_type
                            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                            ON CONFLICT(instrument_id) DO UPDATE SET
                                title = CASE WHEN Projects.title IS NULL OR Projects.title='' THEN NULLIF(excluded.title,'') ELSE Projects.title END,
                                agreement_structure = CASE WHEN Projects.agreement_structure IS NULL OR Projects.agreement_structure='' THEN NULLIF(excluded.agreement_structure,'') ELSE Projects.agreement_structure END,
                                creditor = CASE WHEN Projects.creditor IS NULL OR Projects.creditor='' THEN NULLIF(excluded.creditor,'') ELSE Projects.creditor END,
                                agreement_date = CASE WHEN Projects.agreement_date IS NULL OR Projects.agreement_date='' THEN NULLIF(excluded.agreement_date,'') ELSE Projects.agreement_date END,
                                maturity_date = CASE WHEN Projects.maturity_date IS NULL OR Projects.maturity_date='' THEN NULLIF(excluded.maturity_date,'') ELSE Projects.maturity_date END,
                                amount = COALESCE(excluded.amount, Projects.amount),
                                revised_amount = COALESCE(excluded.revised_amount, Projects.revised_amount),
                                economic_sector = CASE WHEN Projects.economic_sector IS NULL OR Projects.economic_sector='' THEN NULLIF(excluded.economic_sector,'') ELSE Projects.economic_sector END,
                                instrument_fund_use = CASE WHEN Projects.instrument_fund_use IS NULL OR Projects.instrument_fund_use='' THEN NULLIF(excluded.instrument_fund_use,'') ELSE Projects.instrument_fund_use END,
                                main_implementing_agency = CASE WHEN Projects.main_implementing_agency IS NULL OR Projects.main_implementing_agency='' THEN NULLIF(excluded.main_implementing_agency,'') ELSE Projects.main_implementing_agency END,
                                currency = COALESCE(excluded.currency, Projects.currency),
                                upload_id = excluded.upload_id,
                                file_type = excluded.file_type
                        '''

                        # Disbursements: keep existing OR overwrite
                        disb_sql_keep = '''
                            INSERT INTO Disbursements (instrument_id, year, amount, upload_id, file_type)
                            VALUES (?,?,?,?,?)
                            ON CONFLICT(instrument_id, year) DO NOTHING
                        '''
                        disb_sql_overwrite = '''
                            INSERT INTO Disbursements (instrument_id, year, amount, upload_id, file_type)
                            VALUES (?,?,?,?,?)
                            ON CONFLICT(instrument_id, year) DO UPDATE SET
                                amount=excluded.amount,
                                upload_id=excluded.upload_id,
                                file_type=excluded.file_type
                        '''

                        project_sql = project_sql_overwrite if overwrite_amounts else project_sql_keep
                        disb_sql    = disb_sql_overwrite    if overwrite_amounts else disb_sql_keep

                        with get_conn() as conn:
                            cur = conn.cursor()
                            cur.execute(
                                "INSERT INTO Upload_Batches "
                                "(upload_id,filename,upload_date,uploaded_by,currency,record_count,file_type) "
                                "VALUES (?,?,?,?,?,?,?)",
                                (
                                    batch_id, uploaded.name,
                                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    st.session_state.username, user_currency,
                                    len(df_projects), detected_fmt
                                )
                            )
                            if proj_tuples:
                                cur.executemany(project_sql, proj_tuples)
                            if disb_tuples:
                                cur.executemany(disb_sql, disb_tuples)
                            conn.commit()

                        mode_label = "overwrite" if overwrite_amounts else "null-fill"
                        log_action(
                            st.session_state.username, "UPLOAD",
                            f"{len(proj_tuples)} projects, {len(disb_tuples)} disbursements "
                            f"(Batch: {batch_id}, mode: {mode_label})"
                        )
                        st.success(
                            f"✅ Saved **{len(proj_tuples)} instruments** "
                            f"and **{len(disb_tuples)} disbursement records** "
                            f"({'amounts overwritten' if overwrite_amounts else 'existing amounts kept'})."
                        )
                        st.balloons()

                    except Exception as e:
                        st.error(f"❌ Error saving: {e}")
                        st.code(traceback.format_exc())

        except Exception as e:
            st.error(f"❌ Could not parse file: {e}")
            st.code(traceback.format_exc())

    # ── Upload History ────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 📂 Upload History")

    batches = pd.read_sql(
        'SELECT upload_id AS "Batch ID", filename AS "File Name", currency AS "Currency",'
        ' record_count AS "Records", upload_date AS "Uploaded On", uploaded_by AS "By"'
        ' FROM Upload_Batches ORDER BY upload_date DESC',
        get_engine()
    )

    if not batches.empty:
        st.dataframe(batches, use_container_width=True, hide_index=True)
        c1, c2 = st.columns([3, 1])
        with c1:
            del_sel = st.selectbox("Select batch to delete",
                                   batches['Batch ID'] + " — " + batches['File Name'])
        with c2:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🚨 Delete Batch", type="primary"):
                b_id = del_sel.split(" — ")[0]
                backup_bytes = export_backup_json()
                with get_conn() as conn2:
                    conn2.execute("DELETE FROM Projects WHERE upload_id=?", (b_id,))
                    conn2.execute("DELETE FROM Disbursements WHERE upload_id=?", (b_id,))
                    conn2.execute("DELETE FROM Upload_Batches WHERE upload_id=?", (b_id,))
                    conn2.commit()
                log_action(st.session_state.username, "DELETE_BATCH", f"{b_id} pre-delete backup exported")
                st.session_state['_last_backup'] = backup_bytes
                st.session_state['_last_backup_name'] = f"backup_before_delete_{b_id}.json"
                st.success(f"Deleted batch {b_id}")
                st.rerun()
    else:
        st.info("No uploads yet.")

    with st.expander("⚠️ Danger Zone: Wipe Entire Database"):
        st.warning("This permanently deletes ALL projects, disbursements, and upload history.")
        if st.button("☠️ Wipe All Data", type="primary"):
            backup_bytes = export_backup_json()
            with get_conn() as conn3:
                conn3.execute("DELETE FROM Projects")
                conn3.execute("DELETE FROM Disbursements")
                conn3.execute("DELETE FROM Upload_Batches")
                conn3.commit()
            log_action(st.session_state.username, "WIPE_DB", "Admin wiped entire database; pre-wipe backup exported")
            st.session_state['_last_backup'] = backup_bytes
            st.session_state['_last_backup_name'] = f"backup_before_wipe_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            st.success("Database cleared.")
            st.rerun()

    if st.session_state.get('_last_backup'):
        st.markdown("---")
        st.download_button(
            "💾 Download pre-delete backup (JSON)",
            data=st.session_state['_last_backup'],
            file_name=st.session_state.get('_last_backup_name', 'backup.json'),
            mime="application/json",
        )


# ─────────────────────────────────────────────
# PAGE: SYSTEM SETTINGS (Admin only)
# ─────────────────────────────────────────────
def page_settings():
    st.title("⚙️ System Settings — Five Year Plans")

    fyp_df = pd.read_sql("SELECT id, name, start_year, end_year FROM FYP_Config ORDER BY start_year", get_engine())

    st.dataframe(fyp_df, use_container_width=True, hide_index=True)
    st.markdown("---")
    st.subheader("Add / Update FYP")

    with st.form("fyp_form"):
        c1, c2, c3 = st.columns([3, 1, 1])
        with c1: fyp_name = st.text_input("FYP Name", placeholder="e.g. 14th FYP (2030-2034)")
        with c2: s_yr = st.number_input("Start Year", min_value=1990, max_value=2100, value=2030)
        with c3: e_yr = st.number_input("End Year",   min_value=1990, max_value=2100, value=2034)
        if st.form_submit_button("💾 Save FYP") and fyp_name:
            conn2 = get_conn()
            conn2.execute('''
                INSERT INTO FYP_Config (name, start_year, end_year) VALUES (?,?,?)
                ON CONFLICT(name) DO UPDATE SET start_year=excluded.start_year, end_year=excluded.end_year
            ''', (fyp_name, s_yr, e_yr))
            conn2.commit(); conn2.close()
            st.success(f"Saved: {fyp_name}")
            st.rerun()


# ─────────────────────────────────────────────
# PAGE: USER MANAGEMENT (Admin only)
# ─────────────────────────────────────────────
def page_users():
    st.title("👥 User Management & Security")

    engine = get_engine()
    users = pd.read_sql(
        "SELECT id, username, role, allowed_partners, view_only, hidden_columns, hidden_projects, hidden_partners FROM Users",
        engine
    )
    try:
        partners_list = pd.read_sql(
            "SELECT DISTINCT creditor FROM Projects WHERE creditor IS NOT NULL ORDER BY creditor",
            engine
        )['creditor'].tolist()
        proj_df = pd.read_sql(
            "SELECT DISTINCT instrument_id, title, creditor FROM Projects ORDER BY instrument_id",
            engine
        )
        project_options = [
            f"{r['instrument_id']} - {r['title'] or 'Unknown'} ({r['creditor'] or 'N/A'})"
            for _, r in proj_df.iterrows()
        ]
    except Exception as e:
        partners_list, project_options = [], []
        log_action(st.session_state.get('username', 'system'), "ERROR",
                   f"page_users: partners/project query failed: {e}")

    disp = users.copy()
    disp['view_only'] = disp['view_only'].map({1: '🔒 View Only', 0: '✏️ Edit'})
    st.dataframe(disp.drop(columns=['id']), use_container_width=True, hide_index=True)

    st.markdown("---")

    col_opts = {
        'agreement_date': 'Agreement Date',
        'maturity_date':  'Maturity Date',
        'amount':         'Original Amount',
        'revised_amount': 'Revised Amount',
    }

    all_usernames = users['username'].tolist()

    if 'user_edit_sel' not in st.session_state:
        st.session_state['user_edit_sel'] = '— select —'

    # Apply any pending reset (queued from the save handler below) BEFORE
    # the selectbox widget is instantiated — Streamlit forbids writing to a
    # widget's session_state key after it has already rendered this run.
    if st.session_state.pop('_reset_user_sel', False):
        st.session_state['user_edit_sel'] = '— select —'

    def _on_user_sel_change():
        for k in ['_uf_role', '_uf_view_only', '_uf_hcols', '_uf_hprojs', '_uf_hpart']:
            st.session_state.pop(k, None)

    st.subheader("🔍 Select User to Edit / Add New")
    inspect_sel = st.selectbox(
        "Select an existing user to edit, or leave at '— select —' to add a new one",
        ["— select —"] + all_usernames,
        key='user_edit_sel',
        on_change=_on_user_sel_change,
    )

    prefill = {
        'username': '', 'role': 'User', 'view_only': True,
        'hidden_cols': [], 'hidden_projs_raw': [], 'hidden_partners': [],
    }

    if inspect_sel != "— select —":
        row = users[users['username'] == inspect_sel].iloc[0]
        prefill['username']         = row['username']
        prefill['role']             = row['role']
        prefill['view_only']        = bool(row['view_only'])
        prefill['hidden_cols']      = [c for c in str(row.get('hidden_columns') or '').split(',') if c]
        stored_ids                  = [p for p in str(row.get('hidden_projects') or '').split(',') if p]
        prefill['hidden_projs_raw'] = stored_ids
        prefill['hidden_partners']  = [p for p in str(row.get('hidden_partners') or '').split(',') if p]

        st.info(
            f"**User:** `{row['username']}`  |  "
            f"**Role:** `{row['role']}`  |  "
            f"**Access:** {'🔒 View Only' if row['view_only'] else '✏️ Edit'}  \n"
            f"**Hidden columns:** `{prefill['hidden_cols'] or 'none'}`  |  "
            f"**Hidden project IDs:** `{stored_ids or 'none'}`  |  "
            f"**Hidden partners:** `{prefill['hidden_partners'] or 'none'}`  \n"
            f"*(Leave the password field blank to keep the current password unchanged.)*"
        )

    st.markdown("---")
    st.subheader("➕ Add / Update User")

    id_to_label = {}
    for opt in project_options:
        pid = opt.split(" - ")[0]
        id_to_label[pid] = opt
    default_proj_labels = [id_to_label[i] for i in prefill['hidden_projs_raw'] if i in id_to_label]

    form_key = f"user_form_{inspect_sel}"

    with st.form(form_key):
        c1, c2 = st.columns(2)
        with c1:
            new_user  = st.text_input("Username *", value=prefill['username'])
            new_pw    = st.text_input(
                "Password (leave blank to keep existing)",
                type="password",
                help="Leave blank when editing an existing user to keep their current password."
            )
            role      = st.selectbox("Role", ["User", "Admin"],
                                     index=0 if prefill['role'] == 'User' else 1)
            view_only = st.checkbox("View Only (cannot upload/edit data)", value=prefill['view_only'])

        with c2:
            hidden_cols = st.multiselect(
                "Hide Columns from This User",
                list(col_opts.keys()),
                default=prefill['hidden_cols'],
                format_func=lambda x: col_opts[x]
            )
            hidden_projs = st.multiselect(
                "Hide Specific Projects from This User",
                project_options,
                default=default_proj_labels
            )
            hidden_partners = st.multiselect(
                "Hide Development Partners from This User",
                partners_list,
                default=[p for p in prefill['hidden_partners'] if p in partners_list],
                help="Selected partners will be completely hidden from this user's dashboard."
            )

        if st.form_submit_button("💾 Save User"):
            if new_user:
                conn2 = get_conn()
                hidden_ids      = ','.join([p.split(" - ")[0] for p in hidden_projs])
                hidden_col_str  = ','.join(hidden_cols)
                hidden_part_str = ','.join(hidden_partners)

                existing = conn2.execute("SELECT password FROM Users WHERE username=?", (new_user,)).fetchone()
                pw_hash  = hash_pw(new_pw) if new_pw else (existing[0] if existing else hash_pw('changeme'))

                conn2.execute('''
                    INSERT INTO Users (username,password,role,allowed_partners,view_only,hidden_columns,hidden_projects,hidden_partners)
                    VALUES (?,?,?,?,?,?,?,?)
                    ON CONFLICT(username) DO UPDATE SET
                        password=excluded.password,
                        role=excluded.role,
                        allowed_partners=excluded.allowed_partners,
                        view_only=excluded.view_only,
                        hidden_columns=excluded.hidden_columns,
                        hidden_projects=excluded.hidden_projects,
                        hidden_partners=excluded.hidden_partners
                ''', (new_user, pw_hash, role, 'All', int(view_only),
                      hidden_col_str, hidden_ids, hidden_part_str))
                conn2.commit(); conn2.close()
                log_action(st.session_state.username, "USER_SAVE",
                           f"Saved user: {new_user}, hidden_partners={hidden_part_str}")
                st.success(f"✅ User '{new_user}' saved.")
                st.session_state['_reset_user_sel'] = True
                st.rerun()

    st.markdown("---")
    st.subheader("Delete User")
    del_users = [u for u in users['username'].tolist() if u != 'admin']
    if del_users:
        del_sel = st.selectbox("Select user to delete", del_users)
        if st.button("🗑️ Delete User", type="secondary"):
            conn3 = get_conn()
            conn3.execute("DELETE FROM Users WHERE username=?", (del_sel,))
            conn3.commit(); conn3.close()
            log_action(st.session_state.username, "USER_DELETE", del_sel)
            st.success(f"Deleted user: {del_sel}")
            st.rerun()
    else:
        st.caption("No non-admin users to delete.")


# ─────────────────────────────────────────────
# PAGE: AUDIT LOG (Admin only)
# ─────────────────────────────────────────────
def page_audit():
    st.title("📋 Audit Log")
    log = pd.read_sql(
        "SELECT username, action, detail, timestamp FROM Audit_Log ORDER BY timestamp DESC LIMIT 500",
        get_engine()
    )
    st.dataframe(log, use_container_width=True, hide_index=True)
    st.caption(f"Showing up to 500 most recent entries. Total: {len(log)}")


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    if not st.session_state.logged_in:
        show_login()
        return

    page = render_sidebar()

    if   page == "Dashboard":                                                    page_dashboard()
    elif page == "Exchange Rates":                                               page_rates()
    elif page == "Upload & Manage Data" and st.session_state.role == 'Admin':   page_upload()
    elif page == "User Management"      and st.session_state.role == 'Admin':   page_users()
    elif page == "System Settings"      and st.session_state.role == 'Admin':   page_settings()
    elif page == "Audit Log"            and st.session_state.role == 'Admin':   page_audit()
    else:
        st.warning("⛔ Access denied. Contact your administrator.")


if __name__ == "__main__":
    main()
