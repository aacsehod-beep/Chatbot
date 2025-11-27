import pandas as pd
from difflib import get_close_matches

# Load Excel file
file_path = "data/faculty_data.xlsx" #r"C:\Users\Administrator\Desktop\AU\Chatbot\faculty_data.xlsx"
xls = pd.ExcelFile(file_path)

sheet_name = next((s for s in xls.sheet_names if "faculty" in s.lower()), None)
if not sheet_name:
    raise FileNotFoundError("❌ No faculty sheet found.")

df = pd.read_excel(file_path, sheet_name=sheet_name, header=1)
df.columns = df.columns.str.strip()


# =========================================================
# Helper: formatted response
# =========================================================
def format_results(rows):
    output = []
    for _, row in rows.iterrows():
        block = (
            f"Name: {row['Name']}\n"
            f"Designation: {row['Designation']}\n"
            f"DOJ: {row['DOJ']}\n"
            f"Department: {row['Department']}\n"
            f"Phone No: {row['Phone no']}\n"
            f"CUG No: {row['cug no']}\n"
            f"Official Mail: {row['Official Mail ID']}\n"
            f"Personal Mail: {row['Personal Mail ID']}\n"
            "-----------------------------"
        )
        output.append(block)
    return "\n".join(output)


# =========================================================
# MAIN FUNCTION
# =========================================================
def get_faculty_info(query):
    query = query.strip().lower()

    # ===== HELP FEATURE =====
    if "help" in query or "sample" in query or "what can" in query:
        return (
            "📌 **You can search using:**\n\n"
            "🔹 Name (supports fuzzy search)\n"
            "   - 'praveen'\n"
            "   - 'pravnn'\n"
            "   - 'swana dept'\n\n"
            "🔹 Phone Number\n"
            "   - '98765'\n"
            "   - 'phone 9123456780'\n\n"
            "🔹 CUG Number\n"
            "   - 'cug no 12345'\n"
            "   - 'swapna cug no'\n\n"
            "🔹 Email Search\n"
            "   - 'gmail'\n"
            "   - 'ravikanth official mail'\n"
            "   - 'swapna email'\n\n"
            "🔹 Designation Search\n"
            "   - 'hod - all hod details will come'\n"
            "   - 'assistant professor'\n"
            "   - 'assistant professor'\n\n"
            "🔹 Department-Wise Listing\n"
            "   - 'school of engineering- all cse list come'\n"
            "   -'school of science- all cse list come'\n"
            "   - 'faculty in school of engineering '\n"
            "   - 'list school of engineering '\n\n"
            "🟢 Ask anything like:\n"
            "  - 'swapna official mail'\n"
            "  - 'faculty personal mail'\n"
        )

    # Column references
    name_col = 'Name'
    phone_col = 'Phone no'
    cug_col = 'cug no'
    dept_col = 'Department'
    official_mail_col = 'Official Mail ID'
    personal_mail_col = 'Personal Mail ID'
    designation_col = 'Designation'

    # ======================================================
    # 1️⃣ Detect field keywords (multi-word supported)
    # ======================================================
    FIELD_KEYWORDS = {
        "cug no": cug_col,
        "cug number": cug_col,
        "cug": cug_col,

        "phone number": phone_col,
        "mobile number": phone_col,
        "phone no": phone_col,
        "phone": phone_col,
        "mobile": phone_col,
        "contact": phone_col,

        "official mail id": official_mail_col,
        "official mail": official_mail_col,
        "official email": official_mail_col,

        "personal mail id": personal_mail_col,
        "personal mail": personal_mail_col,
        "personal email": personal_mail_col,

        "designation": designation_col,
        "hod": designation_col,
        "professor": designation_col,
        "assistant professor": designation_col,
        "associate professor": designation_col,
        "lab assistant": designation_col,

        "department": dept_col,
        "dept": dept_col
    }

    selected_column = None
    for key in FIELD_KEYWORDS:
        if key in query:
            selected_column = FIELD_KEYWORDS[key]
            break

    # ======================================================
    # 2️⃣ Search by numeric input (phone / cug)
    # ======================================================
    numeric = ''.join([c for c in query if c.isdigit()])
    if numeric:
        rows = df[
            df[phone_col].astype(str).str.contains(numeric, na=False) |
            df[cug_col].astype(str).str.contains(numeric, na=False)
        ]
        if not rows.empty:
            return format_results(rows)

    # ======================================================
    # 3️⃣ Email-Based Search
    # ======================================================
    if "mail" in query or "email" in query or "@" in query:
        rows = df[
            df[official_mail_col].astype(str).str.contains(query.replace(" ", ""), case=False, na=False) |
            df[personal_mail_col].astype(str).str.contains(query.replace(" ", ""), case=False, na=False)
        ]
        if not rows.empty:
            return format_results(rows)

    # ======================================================
    # 4️⃣ Designation search
    # ======================================================
    if "professor" in query or "hod" in query or "assistant" in query:
        rows = df[df[designation_col].str.lower().str.contains(query, na=False)]
        if not rows.empty:
            return format_results(rows)

    # ======================================================
    # 5️⃣ Department-wise Listing
    # ======================================================
    departments = df[dept_col].dropna().unique().tolist()
    dept_lower = [d.lower() for d in departments]

    probable_dept = get_close_matches(query, dept_lower, n=1, cutoff=0.5)

    if probable_dept:
        matched_dept = probable_dept[0]
        rows = df[df[dept_col].str.lower() == matched_dept]
        if not rows.empty:
            return f"📌 Faculty in {matched_dept.upper()} Department:\n\n" + format_results(rows)

    # Also match partial words: "cse dept", "ece department"
    for dept in dept_lower:
        if dept in query:
            rows = df[df[dept_col].str.lower() == dept]
            if not rows.empty:
                return f"📌 Faculty in {dept.upper()} Department:\n\n" + format_results(rows)

    # ======================================================
    # 6️⃣ Fuzzy Name Search
    # ======================================================
    all_names = df[name_col].astype(str).str.lower().tolist()
    words = query.split()

    probable_name = get_close_matches(words[0], all_names, n=1, cutoff=0.5)

    if probable_name:
        rows = df[df[name_col].str.lower().str.contains(probable_name[0])]
    else:
        rows = df[df[name_col].str.lower().str.contains(words[0])]

    if rows.empty:
        return f"❌ No faculty found for '{query}'."

    # ======================================================
    # 7️⃣ If specific field requested
    # ======================================================
    if selected_column:
        output = []
        for _, row in rows.iterrows():
            output.append(f"{row[name_col]} — {selected_column}: {row[selected_column]}")
        return "\n".join(output)

    # ======================================================
    # 8️⃣ Default: Full details
    # ======================================================
    return format_results(rows)
