import io
import os
import pandas as pd
import app as dashboard_app
from app import load_raw, detect_format, parse_upload_file


def make_csv_buffer(csv_text, name='test.csv'):
    buf = io.StringIO(csv_text)
    buf.name = name
    return buf


def test_detect_format_flat_csv():
    csv_data = (
        "Instrument Id,Title,Creditor Name,Agreement Structure,Agreement Date,Maturity Date,Amount,Revised Amount,2024,2025\n"
        "ID100,Project A,Partner X,Grant,2024-01-01,2028-12-31,100000,120000,50000,50000\n"
    )
    buf = make_csv_buffer(csv_data)
    df_raw = load_raw(buf)
    fmt, header_row, col_map, year_cols = detect_format(df_raw)

    assert fmt == 'flat'
    assert header_row == 0
    assert 'Instrument Id' in col_map
    assert 2024 in year_cols and 2025 in year_cols


def test_parse_flat_csv_roundtrip():
    csv_data = (
        "Instrument Id,Title,Creditor Name,Agreement Structure,Agreement Date,Maturity Date,Amount,Revised Amount,2024,2025\n"
        "ID100,Project A,Partner X,Grant,2024-01-01,2028-12-31,100000,120000,50000,50000\n"
    )
    buf = make_csv_buffer(csv_data)
    df_projects, df_disbursements, fmt, year_cols = parse_upload_file(buf, user_currency='USD')

    assert fmt == 'flat'
    assert not df_projects.empty
    assert df_projects.loc[0, 'instrument_id'] == 'ID100'
    assert df_projects.loc[0, 'currency'] == 'USD'
    assert len(df_disbursements) == 2
    assert set(df_disbursements['year']) == {2024, 2025}


def test_parse_flat_with_missing_columns():
    csv_data = (
        "Instrument Id,Main Instrument Title,Creditor,Amount,2024\n"
        "ID200,Project B,Partner Y,75000,75000\n"
    )
    buf = make_csv_buffer(csv_data)
    df_projects, df_disbursements, fmt, year_cols = parse_upload_file(buf, user_currency='USD')

    assert fmt == 'flat'
    assert df_projects.loc[0, 'title'] == 'Project B'
    assert df_projects.loc[0, 'creditor'] == 'Partner Y'
    assert df_disbursements.loc[0, 'amount'] == 75000.0


def test_database_upload_stores_file_type(tmp_path):
    import app as dashboard_app
    csv_data = (
        "Instrument Id,Title,Creditor Name,Agreement Structure,Agreement Date,Maturity Date,Amount,Revised Amount,2024\n"
        "ID300,Project C,Partner Z,Loan,2024-02-01,2029-01-31,250000,260000,100000\n"
    )
    buf = make_csv_buffer(csv_data)
    dashboard_app.DB = str(tmp_path / 'test_upload.db')
    dashboard_app.init_db()

    df_projects, df_disbursements, fmt, year_cols = parse_upload_file(buf, user_currency='EUR')
    batch_id = 'BATCH_TEST_1'

    proj_tuples = [(
        r.get('instrument_id'), r.get('title'), r.get('agreement_structure'), r.get('creditor'),
        r.get('agreement_date'), r.get('maturity_date'), r.get('amount'), r.get('revised_amount'),
        r.get('economic_sector'), r.get('instrument_fund_use'), r.get('main_implementing_agency'),
        r.get('currency') or 'EUR', batch_id, fmt
    ) for _, r in df_projects.iterrows()]
    disb_tuples = [(
        r['instrument_id'], int(r['year']), float(r['amount']), batch_id, fmt)
        for _, r in df_disbursements.iterrows()
    ]

    project_sql = '''
        INSERT INTO Projects (
            instrument_id, title, agreement_structure, creditor,
            agreement_date, maturity_date, amount, revised_amount,
            economic_sector, instrument_fund_use, main_implementing_agency, currency, upload_id, file_type
        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    '''
    disb_sql = '''
        INSERT INTO Disbursements (instrument_id, year, amount, upload_id, file_type)
        VALUES (?,?,?,?,?)
    '''

    with dashboard_app.get_conn() as conn:
        conn.execute("INSERT INTO Upload_Batches (upload_id,filename,upload_date,uploaded_by,currency,record_count,file_type) VALUES (?,?,?,?,?,?,?)",
                     (batch_id, 'test_upload.csv', '2026-01-01 00:00:00', 'tester', 'EUR', len(df_projects), fmt))
        conn.executemany(project_sql, proj_tuples)
        conn.executemany(disb_sql, disb_tuples)
        conn.commit()

    with dashboard_app.get_conn() as conn:
        proj_df = pd.read_sql("SELECT instrument_id, file_type FROM Projects", conn)
        disb_df = pd.read_sql("SELECT instrument_id, year, file_type FROM Disbursements", conn)

    assert proj_df.loc[0, 'file_type'] == fmt
    assert all(disb_df['file_type'] == fmt)

def test_backup_db_creates_backup_file(tmp_path):
    dashboard_app.DB = str(tmp_path / 'backup_test.db')
    dashboard_app.init_db()
    backup_path = dashboard_app.backup_db()
    assert backup_path is not None
    assert os.path.exists(backup_path)
    assert os.path.basename(backup_path).startswith('backup_test_')
    assert os.path.dirname(backup_path).endswith('backups')
