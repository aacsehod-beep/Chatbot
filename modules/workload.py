# import pandas as pd
# import re

# # Path to your workload Excel
# FILE_PATH = r"C:\Users\Administrator\Desktop\AU\Chatbot\workload.xlsx"

# HOURS = ["H1", "H2", "H3", "H4", "Lunch", "H5", "H6", "H7", "H8"]
# DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]

# # ------------------- Core Functions -------------------

# def get_faculty_timetable(faculty_name):
#     xls = pd.ExcelFile(FILE_PATH)
#     sheet_match = None
#     for sheet in xls.sheet_names:
#         if faculty_name.lower() in sheet.lower():
#             sheet_match = sheet
#             break

#     if not sheet_match:
#         return None, f"❌ Faculty '{faculty_name}' not found."

#     df = pd.read_excel(FILE_PATH, sheet_name=sheet_match, header=None)
#     timetable = {}
#     for i, day in enumerate(DAYS, start=1):
#         row_data = df.iloc[i, 1:10].tolist()
#         row_data = ["Free" if (pd.isna(x) or str(x).strip() == "") else str(x) for x in row_data]
#         timetable[day] = dict(zip(HOURS, row_data))

#     return timetable, sheet_match


# def get_free_slots(timetable, day_filter=None):
#     free_slots = {}
#     for day, slots in timetable.items():
#         if day_filter and day.lower() != day_filter.lower():
#             continue
#         day_free = [hour for hour, subject in slots.items() if subject == "Free" and hour != "Lunch"]
#         if day_free:
#             free_slots[day] = day_free
#     return free_slots


# def who_is_free(day, hour):
#     xls = pd.ExcelFile(FILE_PATH)
#     free_faculty = []

#     for sheet in xls.sheet_names:
#         timetable, _ = get_faculty_timetable(sheet)
#         if timetable and timetable[day][hour] == "Free":
#             free_faculty.append(sheet)

#     return free_faculty

# # ------------------- Query Parsing -------------------

# def parse_workload_query(query):
#     query_lower = query.lower()
#     xls = pd.ExcelFile(FILE_PATH)

#     # Faculty detection
#     faculty_name = None
#     for sheet in xls.sheet_names:
#         if sheet.lower() in query_lower:
#             faculty_name = sheet
#             break
#         # Partial match
#         if any(word in sheet.lower() for word in query_lower.split()):
#             faculty_name = sheet
#             break

#     # Day detection
#     day_filter = None
#     for d in DAYS:
#         if d.lower() in query_lower:
#             day_filter = d
#             break

#     # Hour detection
#     hour_filter = None
#     for h in HOURS:
#         if h.lower() in query_lower:
#             hour_filter = h
#             break

#     return faculty_name, day_filter, hour_filter

# # ------------------- GUI-Friendly Function -------------------

# def gui_get_workload(query):
#     faculty_name, day_filter, hour_filter = parse_workload_query(query)

#     # Reverse query → "who is free on Monday H2"
#     if "who" in query.lower() and "free" in query.lower() and day_filter and hour_filter:
#         free_faculty = who_is_free(day_filter, hour_filter)
#         if free_faculty:
#             return f"🕒 Faculty free on {day_filter} {hour_filter}: {', '.join(free_faculty)}"
#         else:
#             return f"❌ No one is free on {day_filter} {hour_filter}."

#     if not faculty_name:
#         return "❌ Please mention a faculty name (e.g., 'Aruna timetable') or ask 'who is free on Monday H2'."

#     timetable, result = get_faculty_timetable(faculty_name)
#     if timetable is None:
#         return result

#     if "free" in query.lower():
#         free_slots = get_free_slots(timetable, day_filter)
#         if free_slots:
#             response = f"🕒 Free slots for {faculty_name}:\n"
#             for day, slots in free_slots.items():
#                 response += f"{day}: {', '.join(slots)}\n"
#             return response.strip()
#         else:
#             if day_filter:
#                 return f"✅ {faculty_name} has no free slots on {day_filter}."
#             else:
#                 return f"✅ {faculty_name} has no free slots."

#     # Full timetable
#     response = f"📅 Timetable for {faculty_name}:\n"
#     for day, slots in timetable.items():
#         response += f"{day}:\n"
#         for hour, subject in slots.items():
#             response += f"  {hour}: {subject}\n"
#         response += "\n"
#     return response.strip()
import pandas as pd
import re
from datetime import datetime, time

FILE_PATH = "data/workload.xlsx" #r"C:\Users\Administrator\Desktop\AU\Chatbot\workload.xlsx"

HOURS = ["H1", "H2", "H3", "H4", "Lunch", "H5", "H6", "H7", "H8"]
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]

_excel_cache = None


# ------------------- EXCEL LOADER -------------------

def load_excel():
    global _excel_cache
    if _excel_cache is None:
        _excel_cache = pd.ExcelFile(FILE_PATH)
    return _excel_cache


# ------------------- BASIC HELPERS -------------------

def find_best_faculty_match(query):
    xls = load_excel()
    query = query.lower()

    for sheet in xls.sheet_names:
        if sheet.lower() == query:
            return sheet

    for sheet in xls.sheet_names:
        if sheet.lower() in query:
            return sheet

    query_words = query.split()
    for sheet in xls.sheet_names:
        if any(word in sheet.lower() for word in query_words):
            return sheet

    return None


def get_faculty_timetable(faculty_name):
    xls = load_excel()
    sheet = find_best_faculty_match(faculty_name)

    if not sheet:
        return None, f"❌ Faculty '{faculty_name}' not found."

    df = pd.read_excel(FILE_PATH, sheet_name=sheet, header=None)

    timetable = {}
    for i, day in enumerate(DAYS, start=1):
        row_data = df.iloc[i, 1:10].tolist()
        cleaned = ["Free" if (pd.isna(x) or str(x).strip() == "") else str(x) for x in row_data]
        timetable[day] = dict(zip(HOURS, cleaned))

    return timetable, sheet


def get_free_slots(timetable, day_filter=None):
    free = {}
    for day, slots in timetable.items():
        if day_filter and day.lower() != day_filter.lower():
            continue

        empty = [h for h, v in slots.items() if v == "Free" and h != "Lunch"]
        if empty:
            free[day] = empty
    return free


# ------------------- NEW FEATURE 1 -------------------
#   CHECK WHO IS FREE RIGHT NOW
# ----------------------------------------------------

def get_current_hour():
    """Map current time to timetable hour."""
    now = datetime.now().time()

    # You can modify these mappings to match your institution timings
    hour_map = {
        "H1": (time(9, 0), time(10, 0)),
        "H2": (time(10, 0), time(11, 0)),
        "H3": (time(11, 0), time(12, 0)),
        "H4": (time(12, 0), time(13, 0)),
        "Lunch": (time(13, 0), time(14, 0)),
        "H5": (time(14, 0), time(15, 0)),
        "H6": (time(15, 0), time(16, 0)),
        "H7": (time(16, 0), time(17, 0)),
        "H8": (time(17, 0), time(18, 0)),
    }

    for hour, (start, end) in hour_map.items():
        if start <= now < end:
            return hour
    return None


def who_is_free_now():
    """Return faculty free at the current system time."""
    day = datetime.now().strftime("%A")
    if day not in DAYS:
        return "Today is a holiday or weekend."

    current_hour = get_current_hour()
    if not current_hour:
        return "Not within college hours."

    xls = load_excel()
    free_list = []

    for sheet in xls.sheet_names:
        timetable, _ = get_faculty_timetable(sheet)
        if timetable and timetable.get(day, {}).get(current_hour) == "Free":
            free_list.append(sheet)

    if free_list:
        return f"🕒 Free **now** ({day} {current_hour}): {', '.join(free_list)}"
    return f"No faculty is free now ({day} {current_hour})."


# ------------------- NEW FEATURE 2 -------------------
#   WHO IS TAKING <SUBJECT> ON <DAY> <HOUR>
# ----------------------------------------------------

def who_is_teaching(subject, day, hour):
    """Search all faculty for who teaches the specified subject."""
    subject = subject.lower()
    xls = load_excel()
    found = []

    for sheet in xls.sheet_names:
        timetable, _ = get_faculty_timetable(sheet)
        slot = timetable.get(day, {}).get(hour, "")

        if subject in slot.lower():  # partial match
            found.append(sheet)

    return found


# ------------------- QUERY PARSER + MAIN LOGIC -------------------

def parse_workload_query(query):
    q = query.lower()

    faculty = find_best_faculty_match(q)
    day = next((d for d in DAYS if d.lower() in q), None)
    hour = next((h for h in HOURS if h.lower() in q), None)

    # detect subject (word after "taking"/"teaching"/"class")
    subject_match = re.search(r"(taking|teaching|class)\s+([a-zA-Z0-9 ]+)", q)
    subject = subject_match.group(2).strip() if subject_match else None

    return faculty, day, hour, subject


def gui_get_workload(query):
    faculty, day, hour, subject = parse_workload_query(query)
    q = query.lower()

    # ✨ 1. Check who is free NOW
    if "free now" in q or "free right now" in q or "who is free now" in q:
        return who_is_free_now()

    # ✨ 2. Who is taking Python on Monday H2?
    if subject and day and hour:
        teachers = who_is_teaching(subject, day, hour)
        if teachers:
            return f"📘 {subject.title()} on {day} {hour} is taken by: {', '.join(teachers)}"
        else:
            return f"❌ No one is taking {subject.title()} on {day} {hour}."

    # Existing features below...

    # Case: Who is free on Monday H2?
    if "who" in q and "free" in q and day and hour:
        free_list = who_is_free(day, hour)
        if free_list:
            return f"🕒 Faculty free on {day} {hour}: {', '.join(free_list)}"
        return f"❌ No one is free on {day} {hour}."

    if not faculty:
        return "❌ Please mention a faculty name or specify subject/day/hour."

    timetable, sheet = get_faculty_timetable(faculty)
    if timetable is None:
        return sheet

    if "free" in q:
        free = get_free_slots(timetable, day)
        if free:
            msg = f"🕒 Free slots for {sheet}:\n"
            for d, slots in free.items():
                msg += f"{d}: {', '.join(slots)}\n"
            return msg.strip()
        return f"❌ No free hours found for {sheet}."
    # --- Day-specific timetable ---
    if faculty and day and "free" not in q:
        slots = timetable.get(day)
        if slots:
            msg = f"📅 Timetable for {sheet} on {day}:\n"
            for h, s in slots.items():
                msg += f"  {h}: {s}\n"
            return msg.strip()
        else:
            return f"❌ No timetable found for {sheet} on {day}."
    # --- Free slots requested ---
    if "free" in q:
        
        # If a specific day is mentioned
        if day:
            day_slots = timetable.get(day)
            if not day_slots:
                return f"❌ No data for {sheet} on {day}."

            free_hours = [h for h, s in day_slots.items() if s == "Free" and h != "Lunch"]

            if free_hours:
                return f"🕒 Free slots for {sheet} on {day}: {', '.join(free_hours)}"
            else:
                return f"✅ {sheet} has no free slots on {day}."

        # If user says "today"
        if "today" in q:
            today = datetime.now().strftime("%A")
            if today not in DAYS:
                return "❌ Today is not a working day."

            day_slots = timetable.get(today)
            free_hours = [h for h, s in day_slots.items() if s == "Free" and h != "Lunch"]

            if free_hours:
                return f"🕒 Free slots for {sheet} today ({today}): {', '.join(free_hours)}"
            else:
                return f"✅ {sheet} has no free slots today."

        # No specific day → return all days
        free = get_free_slots(timetable)
        if free:
            msg = f"🕒 All free slots for {sheet}:\n"
            for d, slots in free.items():
                msg += f"{d}: {', '.join(slots)}\n"
            return msg.strip()
        else:
            return f"❌ {sheet} has no free slots on any day."

    # Full timetable
    msg = f"📅 Timetable for {sheet}:\n\n"
    for d, slots in timetable.items():
        msg += f"{d}:\n"
        for h, s in slots.items():
            msg += f"  {h}: {s}\n"
        msg += "\n"
    return msg.strip()
