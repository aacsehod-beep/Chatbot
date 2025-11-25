import pandas as pd
import os
import re

# --------------------------- Store weekly attendance data ---------------------------
weekly_data = {}

# --------------------------- Load weekly workbook ---------------------------
def load_week_data(week_name, file_path):
    """Load attendance data for a week and store in weekly_data."""
    if os.path.exists(file_path):
        xls = pd.ExcelFile(file_path)
        sections = {sheet_name: pd.read_excel(xls, sheet_name=sheet_name, header=None)
                    for sheet_name in xls.sheet_names}
        weekly_data[week_name.lower()] = sections
        return f"✅ Loaded {week_name}: {list(sections.keys())}"
    else:
        return f"❌ File not found: {file_path}"

# --------------------------- Find sheet inside a week ---------------------------
def find_sheet(week, section):
    section_clean = re.sub(r"\s+", "", section.lower())
    if week not in weekly_data:
        return None
    for sheet_name, sheet_df in weekly_data[week].items():
        sheet_clean = re.sub(r"\s+", "", sheet_name.lower())
        if sheet_clean == section_clean:
            return sheet_df
    return None

# --------------------------- Find student row ---------------------------
def find_student_row(df, student_identifier):
    start_row = 9
    reg_no_col = 1
    name_col = 2

    student_identifier = str(student_identifier).strip().lower()
    for r in range(start_row, df.shape[0]):
        name = str(df.iloc[r, name_col]).strip().lower()
        regno = str(df.iloc[r, reg_no_col]).strip().lower()
        if student_identifier == regno or student_identifier in name:
            return r, df.iloc[r, name_col], df.iloc[r, reg_no_col]
    return None, None, None

# --------------------------- Get student attendance week-by-week ---------------------------
def get_student_attendance(weeks, section, student_identifier, specific_subject=None):
    if isinstance(weeks, str):
        weeks = [weeks]

    if len(weeks) == 1 and weeks[0].lower() == 'all':
        weeks = list(sorted(weekly_data.keys()))  # keep sorted order

    student_name = None
    student_regno = None
    reports = []

    for week in weeks:
        df = find_sheet(week.lower(), section)
        if df is None:
            reports.append(f"❗ {week}: Section '{section}' not found")
            continue

        subject_row = 7
        last_col = df.shape[1]
        subject_cols = []
        for col in range(4, last_col, 3):
            subject_name = df.iloc[subject_row, col]
            if pd.notna(subject_name) and isinstance(subject_name, str):
                perc_col = col + 2
                subject_cols.append({"subject_col": col, "perc_col": perc_col, "name": subject_name})

        student_row, student_name_tmp, student_regno_tmp = find_student_row(df, student_identifier)
        if student_row is None:
            reports.append(f"❗ {week}: Student '{student_identifier}' not found in section {section}")
            continue

        student_name = student_name_tmp
        student_regno = student_regno_tmp

        report = [f"📊 {week.upper()} Attendance for {student_name} ({student_regno}):"]
        for s in subject_cols:
            subject = s["name"]
            if specific_subject and specific_subject.lower() not in subject.lower():
                continue
            attendance = df.iloc[student_row, s["perc_col"]]
            report.append(f"{subject}: {attendance}")
        reports.append("\n".join(report))

    if not student_name:
        return f"❗ Student '{student_identifier}' not found in section {section}."

    return "\n\n".join(reports)

# --------------------------- Threshold-based attendance ---------------------------
threshold_categories = [
    {"label": ">=75%", "func": lambda x: x >= 75},
    {"label": "65-74%", "func": lambda x: 65 <= x < 75},
    {"label": "50-64%", "func": lambda x: 50 <= x < 65},
    {"label": "<50%", "func": lambda x: x < 50},
    {"label": "<=50%", "func": lambda x: x <= 50}
]

def attendance_threshold_report(threshold_label, section, weeks='all', specific_subject=None):
    # Ensure threshold ends with %
    if "%" not in threshold_label:
        threshold_label += "%"

    section_clean = re.sub(r"\s+", "", section.lower())
    threshold_func = None
    for cat in threshold_categories:
        if cat["label"].lower() == threshold_label.lower():
            threshold_func = cat["func"]
            break
    if not threshold_func:
        return "⚠️ Invalid threshold specified."

    if weeks == 'all':
        week_list = sorted(weekly_data.keys())
    else:
        week_list = [w.lower() for w in weeks]

    # Store student attendance week-wise per subject
    students_dict = {}
    for week in week_list:
        week_data = weekly_data.get(week, {})
        sheet = None
        for sheet_name, df in week_data.items():
            if re.sub(r"\s+", "", sheet_name.lower()) == section_clean:
                sheet = df
                break
        if sheet is None:
            continue

        subject_row = 7
        last_col = sheet.shape[1]
        subject_cols = []
        for col in range(4, last_col, 3):
            subject_name = sheet.iloc[subject_row, col]
            if pd.notna(subject_name) and isinstance(subject_name, str):
                perc_col = col + 2
                subject_cols.append({"subject_col": col, "perc_col": perc_col, "name": subject_name})

        start_row = 9
        for r in range(start_row, sheet.shape[0]):
            student_name = str(sheet.iloc[r, 2]).strip()
            student_regno = str(sheet.iloc[r, 1]).strip()
            if student_regno not in students_dict:
                students_dict[student_regno] = {"name": student_name, "subjects": {}}

            for s in subject_cols:
                att = sheet.iloc[r, s["perc_col"]]
                subject_name = s["name"]
                if specific_subject and specific_subject.lower() not in subject_name.lower():
                    continue  # Skip other subjects
                if pd.notna(att):
                    if subject_name not in students_dict[student_regno]["subjects"]:
                        students_dict[student_regno]["subjects"][subject_name] = {}
                    students_dict[student_regno]["subjects"][subject_name][week] = att

    # Build report
    report = f"📊 Attendance report for {section.upper()} ({threshold_label})\n"
    any_matches = False
    for regno, info in students_dict.items():
        for subject, week_att in info["subjects"].items():
            matched_weeks = {w: v for w, v in week_att.items() if threshold_func(v)}
            if matched_weeks:
                any_matches = True
                report += f"\nStudent: {info['name']} ({regno})\n"
                for w, val in matched_weeks.items():
                    icon = "✅" if threshold_label in [">=75%"] else ("⚠️" if threshold_label in ["65-74%", "50-64%"] else "❌")
                    report += f"  {icon} {subject} ({w.upper()}): {val}%\n"

    if not any_matches:
        report += "\nNo students found for this criteria."

    return report

# --------------------------- GUI-friendly wrapper ---------------------------
def gui_get_attendance(query):
    query = query.strip()

    # Handle threshold queries with optional specific subject
    threshold_patterns = [">=75", "65-74", "50-64", "<50", "<=50"]
    query_lower = query.lower()
    for pat in threshold_patterns:
        if pat in query_lower:
            section_match = re.search(r"[a-zA-Z]{3,}-\d+[a-zA-Z]?", query)
            week_matches = re.findall(r"week\d+", query_lower)
            specific_subject = None
            if section_match:
                after_section = query.split(section_match.group())[-1].strip()
                # If text after section is not week numbers, treat as subject
                after_tokens = after_section.split()
                if after_tokens:
                    for token in after_tokens:
                        if not re.match(r"week\d+", token.lower()):
                            specific_subject = token
                            break
                section = section_match.group()
                weeks = week_matches if week_matches else 'all'
                return attendance_threshold_report(pat, section, weeks, specific_subject)
            else:
                return "⚠️ Please provide a valid section (e.g., CSE-2A) for threshold query."

    # Normal student attendance queries
    tokens = query.split()
    if 'attendance' not in [t.lower() for t in tokens]:
        return "⚠️ Please include the word 'attendance' in your query."

    section = None
    weeks = []
    student_identifier = None
    subject = None

    for t in tokens:
        if re.match(r"week\d+", t.lower()):
            weeks.append(t.lower())
        elif re.match(r"[a-zA-Z]{3,}-\d+[a-zA-Z]?", t):
            section = t
        elif t.lower() != 'attendance':
            if not student_identifier:
                student_identifier = t
            elif not subject:
                subject = t

    if not section or not student_identifier:
        return "⚠️ Provide both section and student name or RegNo."

    if not weeks:
        weeks = ['all']

    return get_student_attendance(weeks, section, student_identifier, subject)

# --------------------------- Example: preload weeks ---------------------------
week_files = {
    "week1": r"C:\Users\Administrator\Desktop\AU\Chatbot\week1.xlsx",
    "week2": r"C:\Users\Administrator\Desktop\AU\Chatbot\week2.xlsx"
}

for week, path in week_files.items():
    load_week_data(week, path)
