"""
DCDMD Dashboard - USER READ-ONLY VERSION
Dashboard-only interface for 3rd-party users and staff
No upload, management, or settings pages

IMPORTANT: this file now shares the exact same Postgres/Supabase connection,
password verification, and hiding logic as app.py. It previously ran on a
disconnected local SQLite file (dashboard_data.db) with plain SHA-256
passwords, which is why data here could look stale/wrong and why
per-user hiding configured in app.py's User Management page had no effect.
"""
import os
import streamlit as st
import pandas as pd
import numpy as np
import hashlib
import hmac
import secrets
import re
import traceback
from urllib.parse import quote_plus
from sqlalchemy import create_engine
import plotly.express as px
from datetime import datetime, date, time as dt_time
from pathlib import Path

# ─────────────────────────────────────────────
# OPENPYXL BUG FIX (kept for parity — unused here since this app never
# parses uploads, but harmless and keeps the two files easy to diff)
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
# DATABASE — same Postgres (Supabase) engine + sqlite3-compatible wrapper
# as app.py. This app is READ-ONLY: it never calls init_db(), never
# creates tables, and never writes anything except Audit_Log.
# ─────────────────────────────────────────────
@st.cache_resource(show_spinner=False)
def get_engine():
    cfg = st.secrets["postgres"]
    url = (
        f"postgresql+psycopg2://{cfg['user']}:{quote_plus(str(cfg['password']))}"
        f"@{cfg['host']}:{cfg['port']}/{cfg['dbname']}"
    )
    return create_engine(url, pool_pre_ping=True, pool_size=5, max_overflow=10)


class _CursorWrapper:
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
    return PGConn(get_engine().raw_connection())


# ─────────────────────────────────────────────
# HELPERS (password verification must match app.py exactly, or accounts
# created/upgraded there won't be able to log in here)
# ─────────────────────────────────────────────
def hash_pw(pw, iterations=260_000):
    salt = secrets.token_hex(16)
    dk = hashlib.pbkdf2_hmac('sha256', str(pw).encode(), bytes.fromhex(salt), iterations)
    return f"pbkdf2_sha256${iterations}${salt}${dk.hex()}"


def verify_pw(pw, stored):
    """Returns (matches, needs_upgrade). Accepts both the new salted format and
    the old unsalted sha256 hex digest, so existing accounts keep working."""
    if not stored:
        return False, False
    if stored.startswith('pbkdf2_sha256$'):
        try:
            _, iter_s, salt, hash_hex = stored.split('$')
            dk = hashlib.pbkdf2_hmac('sha256', str(pw).encode(), bytes.fromhex(salt), int(iter_s))
            return hmac.compare_digest(dk.hex(), hash_hex), False
        except Exception:
            return False, False
    else:
        legacy_match = hmac.compare_digest(hashlib.sha256(str(pw).encode()).hexdigest(), stored)
        return legacy_match, legacy_match


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

# Tokens that mean "no value" wherever they show up — as literal text stored
# in the DB from older uploads, not just genuinely-empty cells.
BLANK_TOKENS = ('', 'nan', 'none', 'n/a', '-', '#ref!', 'nat')


def clean_display(v, blank='N/A'):
    """Render-time guard: turns literal junk text ('NaN', 'None', ...) that may
    already be sitting in the database into a clean placeholder for display."""
    if v is None:
        return blank
    try:
        if pd.isna(v):
            return blank
    except (TypeError, ValueError):
        pass
    s = str(v).strip()
    return blank if s.lower() in BLANK_TOKENS else s


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
        st.markdown("**User Portal — Read-Only Access**")
        st.markdown("---")
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        if st.button("Login", use_container_width=True, type="primary"):
            if not username or not password:
                return st.error("Enter both username and password.")
            conn = get_conn()
            row = conn.execute(
                "SELECT * FROM Users WHERE username=?",
                (username,)
            ).fetchone()
            conn.close()
            matches, needs_upgrade = verify_pw(password, row[2]) if row else (False, False)
            if row and matches:
                if needs_upgrade:
                    with get_conn() as uconn:
                        uconn.execute("UPDATE Users SET password=? WHERE username=?",
                                      (hash_pw(password), username))
                        uconn.commit()
                st.session_state.logged_in       = True
                st.session_state.username        = row[1]
                st.session_state.role            = row[3]
                st.session_state.allowed_partners = row[4]
                st.session_state.view_only       = bool(row[5])
                st.session_state.hidden_cols     = [c for c in row[6].split(',') if c] if row[6] else []
                st.session_state.hidden_projs    = [p for p in row[7].split(',') if p] if row[7] else []
                # Partner names can contain commas, so this list uses '|' as its delimiter
                st.session_state.hidden_partners = [p for p in row[8].split('|') if p] if (len(row) > 8 and row[8]) else []
                log_action(username, "LOGIN_USER")
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
    st.sidebar.markdown("**📊 Dashboard** (Read-Only)")
    st.sidebar.markdown("---")
    if st.sidebar.button("🚪 Logout"):
        log_action(st.session_state.username, "LOGOUT_USER")
        for k, v in _defaults.items(): st.session_state[k] = v
        st.rerun()


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
        return st.warning("⚠️ No data available yet.")

    # ── Security filters ────────────────────────────────────────────────────
    if st.session_state.hidden_projs:
        df_all = df_all[~df_all['instrument_id'].isin(st.session_state.hidden_projs)]

    if st.session_state.allowed_partners != 'All':
        allowed = [p.strip() for p in st.session_state.allowed_partners.split('|')]
        df_all  = df_all[df_all['creditor'].isin(allowed)]

    hidden_partners = st.session_state.get('hidden_partners') or []
    if hidden_partners:
        hidden_set = {str(hp).strip().lower() for hp in hidden_partners if hp}
        def _is_hidden_partner(creditor_val):
            if creditor_val is None or (isinstance(creditor_val, float) and pd.isna(creditor_val)):
                return False
            return str(creditor_val).strip().lower() in hidden_set
        df_all = df_all[~df_all['creditor'].apply(_is_hidden_partner)]

    if df_all.empty:
        return st.warning("No data available for your account.")

    # ── Sidebar filters ─────────────────────────────────────────────────────
    st.sidebar.markdown("### 🔍 Filters")
    creditors  = ['All'] + sorted(df_all['creditor'].dropna().unique().tolist())
    agencies   = ['All'] + sorted(df_all['main_implementing_agency'].dropna().unique().tolist())
    structures = ['All'] + sorted(df_all['agreement_structure'].dropna().unique().tolist())

    sel_cred    = st.sidebar.selectbox("Development Partner", creditors)
    sel_agency  = st.sidebar.selectbox("Implementing Agency", agencies)
    sel_struct  = st.sidebar.selectbox("Instrument Type", structures)

    # ── Currency mode ───────────────────────────────────────────────────────
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

    # ── Apply filters ───────────────────────────────────────────────────────
    dff = df_all.copy()
    if sel_cred   != 'All': dff = dff[dff['creditor'] == sel_cred]
    if sel_agency != 'All': dff = dff[dff['main_implementing_agency'] == sel_agency]
    if sel_struct != 'All': dff = dff[dff['agreement_structure'] == sel_struct]

    # ── Currency conversion ─────────────────────────────────────────────────
    if currency_mode == "Aggregate (BTN / Nu.)":
        missing = [c for c in dff['currency'].dropna().unique()
                   if str(c).upper() not in rates_df.index]
        if missing:
            st.warning(f"⚠️ Missing rates for: **{', '.join(missing)}** — shown as 1:1.")

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

    # ── KPI Cards ───────────────────────────────────────────────────────────
    total   = dff['display_amount'].sum()
    grants  = dff[dff['type'] == 'Grant']['display_amount'].sum()
    loans   = dff[dff['type'] == 'Loan']['display_amount'].sum()
    n_proj  = dff['instrument_id'].nunique()
    n_part  = 1 if sel_cred != 'All' else dff['creditor'].nunique()

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("💰 Total Amount",  f"{total:,.2f}",  unit)
    k2.metric("🤝 Total Grants",     f"{grants:,.2f}", unit)
    k3.metric("🏦 Total Loans",      f"{loans:,.2f}",  unit)
    k4.metric("📁 Projects",         f"{n_proj}")
    k5.metric("🌐 Dev Partners",     f"{n_part}")
    st.markdown("---")

    color_map = {
        'Grant': COLORS['grant'],
        'Loan':  COLORS['loan'],
        'Other': COLORS['other'],
    }

    # ── Row 1: FYP Bar + Yearly Timeline ────────────────────────────────────
    col1, col2 = st.columns([1, 1.2])

    with col1:
        st.subheader("📋 Amount for Five-Year Plan")
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
        fyp_opts = ["Show All"] + sorted([
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

    # ── Row 2: Creditor Donut + Agency Bar ──────────────────────────────────
    col3, col4 = st.columns([1, 1.2])

    with col3:
        st.subheader("🌐 Creditor Share")
        cred_pie = (dff[dff['display_amount'] > 0]
                    .groupby('creditor')['display_amount'].sum().reset_index())
        if not cred_pie.empty:
            fig3 = px.pie(cred_pie, values='display_amount', names='creditor',
                          hole=0.42, color_discrete_sequence=px.colors.qualitative.Pastel)
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

    # ── Project Drill-Down ──────────────────────────────────────────────────
    st.subheader("🔍 Project Drill-Down")
    uniq = dff.drop_duplicates('instrument_id').copy()
    uniq['label'] = uniq['instrument_id'] + "  —  " + uniq['title'].fillna('Unknown')
    sel_label = st.selectbox("Search Project", ["— Select a project —"] + sorted(uniq['label'].tolist()))

    if sel_label != "— Select a project —":
        sel_id    = sel_label.split("  —  ")[0].strip()
        proj_rows = dff[dff['instrument_id'] == sel_id]
        p         = proj_rows.iloc[0]
        total_disb = proj_rows['display_amount'].sum()
        ptype = p['type']
        tc = COLORS['grant'] if ptype == 'Grant' else (COLORS['loan'] if ptype == 'Loan' else COLORS['other'])

        def safe_val(col_key):
            if col_key in st.session_state.hidden_cols:
                return "🔒 Hidden"
            return clean_display(p.get(col_key))

        ci1, ci2 = st.columns([1, 1])
        with ci1:
            st.markdown(f"""
            <div style="background:{COLORS['card']};border-radius:12px;padding:20px;border-left:4px solid {tc}">
                <h4 style="margin:0 0 12px;color:#fff">{clean_display(p.get('title'))}</h4>
                <table style="width:100%;color:#ccc;font-size:13px;border-collapse:collapse">
                    <tr><td><b>🆔 Instrument ID</b></td><td style="text-align:right">{p['instrument_id']}</td></tr>
                    <tr><td><b>🏷️ Type</b></td><td style="text-align:right"><span style="color:{tc};font-weight:bold">{ptype}</span></td></tr>
                    <tr><td><b>🤝 Creditor</b></td><td style="text-align:right">{clean_display(p.get('creditor'))}</td></tr>
                    <tr><td><b>🏛️ Agency</b></td><td style="text-align:right">{clean_display(p.get('main_implementing_agency'))}</td></tr>
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

    # ── Data Matrix Table ───────────────────────────────────────────────────
    with st.expander("📄 Full Data Matrix", expanded=False):
        search = st.text_input("🔍 Search (ID, Title, Creditor, Agency):", "")

        base_cols = ['instrument_id', 'title', 'creditor', 'agreement_structure',
                     'main_implementing_agency', 'currency']
        secure_cols = ['agreement_date', 'maturity_date', 'amount', 'revised_amount']
        for sc in secure_cols:
            if sc not in st.session_state.hidden_cols:
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

        rename_dict  = {}
        year_col_list = []
        for c in df_display.columns:
            cs = str(c)
            if re.fullmatch(r'\d{4}(\.0)?', cs.strip()):
                yr = str(int(float(cs)))
                rename_dict[c] = yr
                year_col_list.append(yr)
        df_display = df_display.rename(columns=rename_dict)
        year_col_list = sorted(set(year_col_list))

        if year_col_list:
            df_display[year_col_list] = df_display[year_col_list].fillna(0)
            df_display['Total Disbursed'] = df_display[year_col_list].sum(axis=1)

        if search:
            mask = df_display.astype(str).apply(
                lambda x: x.str.contains(search, case=False, na=False)
            ).any(axis=1)
            df_display = df_display[mask]

        df_display_fmt = df_display.copy()
        for txt_col in ['title', 'creditor', 'agreement_structure', 'main_implementing_agency',
                        'currency', 'agreement_date', 'maturity_date']:
            if txt_col in df_display_fmt.columns:
                df_display_fmt[txt_col] = df_display_fmt[txt_col].apply(lambda v: clean_display(v, blank=''))

        st.dataframe(df_display_fmt, use_container_width=True, hide_index=True)
        st.caption(f"Showing {len(df_display)} projects | Unit: {unit}")


def main():
    if not st.session_state.logged_in:
        show_login()
        return

    render_sidebar()
    page_dashboard()


if __name__ == "__main__":
    main()
