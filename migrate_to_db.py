"""
migrate_to_db.py
One-time migration: Excel files → SQLite database (university.db)

Tables created:
  - faculty          (from faculty_data.xlsx)
  - students         (from student_contacts.xlsx + student_info.xlsx merged)
  - timetable        (from timetable.xlsx — one row per section/day/hour)
  - workload         (from workload.xlsx — one row per faculty/day/hour)
  - attendance       (from week1.xlsx + week2.xlsx — one row per student/subject/week)
"""

import sqlite3
import pandas as pd
import re
import os

DB_PATH = "university.db"

def clean_col(name):
    """Normalize column name to snake_case."""
    name = str(name).strip().lower()
    name = re.sub(r'[^a-z0-9]+', '_', name)
    return name.strip('_')

def create_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn

# ─────────────────────────────────────────────
# 1. FACULTY
# ─────────────────────────────────────────────
def migrate_faculty(conn):
    print("Migrating faculty...")
    df = pd.read_excel("data/faculty_data.xlsx", sheet_name="Faculty-info", header=1)
    df.columns = [clean_col(c) for c in df.columns]
    # Keep only rows that have a name
    df = df[df['name'].notna() & (df['name'].astype(str).str.strip() != '')].copy()
    # Rename for clarity
    df = df.rename(columns={
        's_no': 'sno',
        'phone_no': 'phone',
        'cug_no': 'cug',
        'official_mail_id': 'official_email',
        'personal_mail_id': 'personal_email',
    })
    # Drop columns we don't need
    keep = ['sno', 'name', 'doj', 'designation', 'phone', 'cug', 'role', 'official_email', 'personal_email', 'department']
    df = df[[c for c in keep if c in df.columns]]
    # Normalize phone/cug to string
    for col in ['phone', 'cug']:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: str(int(float(x))) if pd.notna(x) and str(x).replace('.','',1).isdigit() else str(x) if pd.notna(x) else None)

    conn.execute("DROP TABLE IF EXISTS faculty")
    conn.execute("""
        CREATE TABLE faculty (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            sno         TEXT,
            name        TEXT NOT NULL,
            doj         TEXT,
            designation TEXT,
            phone       TEXT,
            cug         TEXT,
            role        TEXT,
            official_email  TEXT,
            personal_email  TEXT,
            department  TEXT
        )
    """)
    df.to_sql("faculty", conn, if_exists="append", index=False)
    print(f"  ✓ faculty: {len(df)} rows")

# ─────────────────────────────────────────────
# 2. STUDENTS
# ─────────────────────────────────────────────
def migrate_students(conn):
    print("Migrating students...")
    # student_contacts has the most complete data; student_info is a near-duplicate
    # We read all sheets from student_contacts, deduplicate by reg_no
    all_rows = []

    xl = pd.ExcelFile("data/student_contacts.xlsx")
    for sheet in xl.sheet_names:
        df = pd.read_excel("data/student_contacts.xlsx", sheet_name=sheet)
        # Some sheets have merged cells or dummy first rows — skip if no real data
        if df.empty:
            continue
        # Detect which column is reg_no
        df.columns = [clean_col(c) for c in df.columns]

        # Normalise column names across sheets
        col_map = {}
        for c in df.columns:
            if 'reg' in c:
                col_map[c] = 'reg_no'
            elif c == 'name_':
                col_map[c] = 'name'
            elif 'student_phone' in c or 'student_contact' in c:
                col_map[c] = 'student_phone'
            elif 'parent_phone' in c or 'parent_contact' in c:
                col_map[c] = 'parent_phone'
            elif 'mail' in c or 'email' in c:
                col_map[c] = 'email'
            elif 'dept' in c:
                col_map[c] = 'dept'
        df = df.rename(columns=col_map)

        # Add section from sheet name if not already a column
        if 'dept' not in df.columns:
            df['dept'] = sheet

        # must have reg_no and name
        if 'reg_no' not in df.columns or 'name' not in df.columns:
            continue

        # Filter valid reg_no rows (pattern like 231U1R1001)
        df = df[df['reg_no'].astype(str).str.match(r'\d{3}U\d[A-Z]\d{4}', na=False)].copy()
        if df.empty:
            continue

        # Keep only key columns
        keep = ['reg_no', 'name', 'student_phone', 'parent_phone', 'email', 'dept']
        df = df[[c for c in keep if c in df.columns]]
        all_rows.append(df)

    if not all_rows:
        print("  ✗ No student rows found")
        return

    combined = pd.concat(all_rows, ignore_index=True)
    # Deduplicate by reg_no keeping first occurrence
    combined = combined.drop_duplicates(subset='reg_no', keep='first')
    # Normalize phones to string
    for col in ['student_phone', 'parent_phone']:
        if col in combined.columns:
            combined[col] = combined[col].apply(
                lambda x: str(x).strip().replace('\n', ', ') if pd.notna(x) else None
            )
    combined['name'] = combined['name'].astype(str).str.strip()
    combined = combined[combined['name'].str.len() > 1]  # remove single-char junk rows

    conn.execute("DROP TABLE IF EXISTS students")
    conn.execute("""
        CREATE TABLE students (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            reg_no        TEXT UNIQUE NOT NULL,
            name          TEXT NOT NULL,
            student_phone TEXT,
            parent_phone  TEXT,
            email         TEXT,
            section       TEXT
        )
    """)
    # rename dept → section for clarity in queries
    if 'dept' in combined.columns:
        combined = combined.rename(columns={'dept': 'section'})

    combined.to_sql("students", conn, if_exists="append", index=False)
    print(f"  ✓ students: {len(combined)} rows")

# ─────────────────────────────────────────────
# 3. TIMETABLE
# ─────────────────────────────────────────────
def migrate_timetable(conn):
    print("Migrating timetable...")
    xl = pd.ExcelFile("data/timetable.xlsx")
    rows = []
    hours = ['H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'H7', 'H8']

    for section in xl.sheet_names:
        raw = pd.read_excel("data/timetable.xlsx", sheet_name=section, header=None)
        # Row 9 (index 9) is often the class incharge info
        class_incharge = None
        if len(raw) > 9:
            ci_val = str(raw.iloc[9, 0]) if pd.notna(raw.iloc[9, 0]) else None
            if ci_val and len(ci_val) > 5:
                class_incharge = ci_val

        # Find the row with H1 in it (the header row)
        header_row = None
        for i, row in raw.iterrows():
            if 'H1' in [str(v) for v in row.values]:
                header_row = i
                break
        if header_row is None:
            continue

        # Data rows follow header_row
        col_map = {v: k for k, v in enumerate(raw.iloc[header_row].values)}
        # days start from header_row+1
        for i in range(header_row + 1, len(raw)):
            row = raw.iloc[i]
            day = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else None
            if not day or day.lower() in ('nan', '', 'none'):
                continue
            # Skip empty rows (all hours NaN)
            hour_values = [str(row.iloc[j]).strip() if pd.notna(row.iloc[j]) else None for j in range(1, 9)]
            if all(v is None or v == 'nan' for v in hour_values):
                continue

            for j, h in enumerate(hours, start=1):
                if j >= len(row):
                    continue
                val = row.iloc[j]
                if pd.isna(val) or str(val).strip() in ('nan', '', 'LUNCH BREAK', 'Lunch Break'):
                    continue
                subject_teacher = str(val).strip().replace('\n', ' ')
                # Try to separate subject and teacher (split on ' - ' or '-')
                subject = subject_teacher
                teacher = None
                if ' - ' in subject_teacher:
                    parts = subject_teacher.split(' - ', 1)
                    subject = parts[0].strip()
                    teacher = parts[1].strip()
                rows.append({
                    'section': section,
                    'day': day.rstrip(),
                    'hour': h,
                    'subject': subject,
                    'teacher': teacher,
                    'class_incharge': class_incharge
                })

    df = pd.DataFrame(rows)
    conn.execute("DROP TABLE IF EXISTS timetable")
    conn.execute("""
        CREATE TABLE timetable (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            section        TEXT NOT NULL,
            day            TEXT NOT NULL,
            hour           TEXT NOT NULL,
            subject        TEXT,
            teacher        TEXT,
            class_incharge TEXT
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_tt_section ON timetable(section)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_tt_day ON timetable(day)")
    df.to_sql("timetable", conn, if_exists="append", index=False)
    print(f"  ✓ timetable: {len(df)} rows across {len(xl.sheet_names)} sections")

# ─────────────────────────────────────────────
# 4. WORKLOAD (faculty schedule)
# ─────────────────────────────────────────────
def migrate_workload(conn):
    print("Migrating faculty workload...")
    xl = pd.ExcelFile("data/workload.xlsx")
    rows = []
    hours = ['H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'H7', 'H8']

    for faculty_name in xl.sheet_names:
        raw = pd.read_excel("data/workload.xlsx", sheet_name=faculty_name, header=None)
        if raw.empty or len(raw.columns) < 2:
            continue

        # First column is Day, rest are H1-H8 (with H0 = LUNCH in middle)
        # Find header row
        header_row = None
        for i, row in raw.iterrows():
            vals = [str(v).strip() for v in row.values]
            if 'H1' in vals:
                header_row = i
                break

        if header_row is None:
            # No header row — assume row 0 is data with columns Day, H1..H8
            raw.columns = ['day', 'H1', 'H2', 'H3', 'H4', 'H0', 'H5', 'H6', 'H7', 'H8'][:len(raw.columns)]
            data_df = raw
        else:
            raw.columns = [str(v).strip() for v in raw.iloc[header_row].values]
            data_df = raw.iloc[header_row + 1:].copy()
            data_df = data_df.rename(columns={data_df.columns[0]: 'day'})

        days_of_week = {'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'}
        for _, row in data_df.iterrows():
            day = str(row.iloc[0]).strip().rstrip() if pd.notna(row.iloc[0]) else None
            if not day or day.lower().strip() not in days_of_week:
                continue
            for h in hours:
                if h not in row.index:
                    continue
                val = row[h]
                if pd.isna(val) or str(val).strip() in ('nan', '', 'LUNCH BREAK', 'Lunch Break'):
                    continue
                content = str(val).strip().replace('\n', ' ')
                rows.append({
                    'faculty': faculty_name.strip(),
                    'day': day.strip(),
                    'hour': h,
                    'subject_section': content
                })

    df = pd.DataFrame(rows)
    conn.execute("DROP TABLE IF EXISTS workload")
    conn.execute("""
        CREATE TABLE workload (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            faculty         TEXT NOT NULL,
            day             TEXT NOT NULL,
            hour            TEXT NOT NULL,
            subject_section TEXT
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_wl_faculty ON workload(faculty)")
    df.to_sql("workload", conn, if_exists="append", index=False)
    print(f"  ✓ workload: {len(df)} rows across {len(xl.sheet_names)} faculty")

# ─────────────────────────────────────────────
# 5. ATTENDANCE
# ─────────────────────────────────────────────
def parse_attendance_sheet(path, week_label, section):
    """
    Parse one attendance sheet.
    Returns list of dicts: {week, section, reg_no, name, subject, held, attended, percentage}
    """
    raw = pd.read_excel(path, sheet_name=section, header=None)
    rows_out = []

    # Find the subject header row (row where 'S.No.' or 'Registration No' appears)
    subject_row = None
    data_start = None
    for i, row in raw.iterrows():
        vals = [str(v).strip() for v in row.values]
        if 'Registration No' in vals or 'Reg.No' in vals:
            subject_row = i
            data_start = i + 2  # skip the Held/Attd/% sub-header row
            break

    if subject_row is None:
        return rows_out

    # Extract subject names from subject_row
    raw_subject_row = list(raw.iloc[subject_row])
    # Subjects appear at columns 4, 7, 10, 13... (every 3rd starting from col 4)
    subjects = []
    col = 4
    while col < len(raw_subject_row):
        subj = str(raw_subject_row[col]).strip()
        if subj and subj != 'nan':
            subjects.append((subj, col))
        col += 3

    # Parse data rows
    for i in range(data_start, len(raw)):
        row = list(raw.iloc[i])
        # Validate: row[0] should be a number (S.No)
        try:
            sno = int(float(str(row[0])))
        except (ValueError, TypeError):
            continue
        reg_no = str(row[1]).strip() if len(row) > 1 and pd.notna(row[1]) else None
        name = str(row[2]).strip() if len(row) > 2 and pd.notna(row[2]) else None
        if not reg_no or not name:
            continue

        for subj_name, start_col in subjects:
            try:
                held = int(float(str(row[start_col]))) if len(row) > start_col and row[start_col] not in (None, '') and str(row[start_col]) != 'nan' else None
                attd = int(float(str(row[start_col + 1]))) if len(row) > start_col + 1 and row[start_col + 1] not in (None, '') and str(row[start_col + 1]) != 'nan' else None
                pct = int(float(str(row[start_col + 2]))) if len(row) > start_col + 2 and row[start_col + 2] not in (None, '') and str(row[start_col + 2]) != 'nan' else None
            except (ValueError, TypeError):
                continue
            rows_out.append({
                'week': week_label,
                'section': section,
                'reg_no': reg_no,
                'name': name,
                'subject': subj_name,
                'held': held,
                'attended': attd,
                'percentage': pct
            })

    return rows_out

def migrate_attendance(conn):
    print("Migrating attendance...")
    all_rows = []
    for week_label, path in [('week1', 'data/week1.xlsx'), ('week2', 'data/week2.xlsx')]:
        xl = pd.ExcelFile(path)
        for section in xl.sheet_names:
            rows = parse_attendance_sheet(path, week_label, section)
            all_rows.extend(rows)

    df = pd.DataFrame(all_rows) if all_rows else pd.DataFrame()
    conn.execute("DROP TABLE IF EXISTS attendance")
    conn.execute("""
        CREATE TABLE attendance (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            week        TEXT NOT NULL,
            section     TEXT NOT NULL,
            reg_no      TEXT NOT NULL,
            name        TEXT,
            subject     TEXT NOT NULL,
            held        INTEGER,
            attended    INTEGER,
            percentage  INTEGER
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_att_reg ON attendance(reg_no)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_att_section ON attendance(section)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_att_week ON attendance(week)")
    if not df.empty:
        df.to_sql("attendance", conn, if_exists="append", index=False)
    print(f"  ✓ attendance: {len(df)} rows")

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    if os.path.exists(DB_PATH):
        os.remove(DB_PATH)
        print(f"Removed existing {DB_PATH}")

    conn = create_connection()
    try:
        migrate_faculty(conn)
        migrate_students(conn)
        migrate_timetable(conn)
        migrate_workload(conn)
        migrate_attendance(conn)
        conn.commit()
        print(f"\n✅ Migration complete → {DB_PATH}")

        # Verification summary
        print("\n── Table row counts ──")
        for table in ['faculty', 'students', 'timetable', 'workload', 'attendance']:
            count = conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
            print(f"  {table}: {count} rows")
    finally:
        conn.close()

if __name__ == "__main__":
    main()
