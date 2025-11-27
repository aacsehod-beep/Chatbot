# # import pandas as pd
# # import re

# # # Path to your workload Excel
# # FILE_PATH = r"C:\Users\Administrator\Desktop\AU\Chatbot\workload.xlsx"

# # HOURS = ["H1", "H2", "H3", "H4", "Lunch", "H5", "H6", "H7", "H8"]
# # DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]

# # # ------------------- Core Functions -------------------

# # def get_faculty_timetable(faculty_name):
# #     xls = pd.ExcelFile(FILE_PATH)
# #     sheet_match = None
# #     for sheet in xls.sheet_names:
# #         if faculty_name.lower() in sheet.lower():
# #             sheet_match = sheet
# #             break

# #     if not sheet_match:
# #         return None, f"❌ Faculty '{faculty_name}' not found."

# #     df = pd.read_excel(FILE_PATH, sheet_name=sheet_match, header=None)
# #     timetable = {}
# #     for i, day in enumerate(DAYS, start=1):
# #         row_data = df.iloc[i, 1:10].tolist()
# #         row_data = ["Free" if (pd.isna(x) or str(x).strip() == "") else str(x) for x in row_data]
# #         timetable[day] = dict(zip(HOURS, row_data))

# #     return timetable, sheet_match


# # def get_free_slots(timetable, day_filter=None):
# #     free_slots = {}
# #     for day, slots in timetable.items():
# #         if day_filter and day.lower() != day_filter.lower():
# #             continue
# #         day_free = [hour for hour, subject in slots.items() if subject == "Free" and hour != "Lunch"]
# #         if day_free:
# #             free_slots[day] = day_free
# #     return free_slots


# # def who_is_free(day, hour):
# #     xls = pd.ExcelFile(FILE_PATH)
# #     free_faculty = []

# #     for sheet in xls.sheet_names:
# #         timetable, _ = get_faculty_timetable(sheet)
# #         if timetable and timetable[day][hour] == "Free":
# #             free_faculty.append(sheet)

# #     return free_faculty

# # # ------------------- Query Parsing -------------------

# # def parse_workload_query(query):
# #     query_lower = query.lower()
# #     xls = pd.ExcelFile(FILE_PATH)

# #     # Faculty detection
# #     faculty_name = None
# #     for sheet in xls.sheet_names:
# #         if sheet.lower() in query_lower:
# #             faculty_name = sheet
# #             break
# #         # Partial match
# #         if any(word in sheet.lower() for word in query_lower.split()):
# #             faculty_name = sheet
# #             break

# #     # Day detection
# #     day_filter = None
# #     for d in DAYS:
# #         if d.lower() in query_lower:
# #             day_filter = d
# #             break

# #     # Hour detection
# #     hour_filter = None
# #     for h in HOURS:
# #         if h.lower() in query_lower:
# #             hour_filter = h
# #             break

# #     return faculty_name, day_filter, hour_filter

# # # ------------------- GUI-Friendly Function -------------------

# # def gui_get_workload(query):
# #     faculty_name, day_filter, hour_filter = parse_workload_query(query)

# #     # Reverse query → "who is free on Monday H2"
# #     if "who" in query.lower() and "free" in query.lower() and day_filter and hour_filter:
# #         free_faculty = who_is_free(day_filter, hour_filter)
# #         if free_faculty:
# #             return f"🕒 Faculty free on {day_filter} {hour_filter}: {', '.join(free_faculty)}"
# #         else:
# #             return f"❌ No one is free on {day_filter} {hour_filter}."

# #     if not faculty_name:
# #         return "❌ Please mention a faculty name (e.g., 'Aruna timetable') or ask 'who is free on Monday H2'."

# #     timetable, result = get_faculty_timetable(faculty_name)
# #     if timetable is None:
# #         return result

# #     if "free" in query.lower():
# #         free_slots = get_free_slots(timetable, day_filter)
# #         if free_slots:
# #             response = f"🕒 Free slots for {faculty_name}:\n"
# #             for day, slots in free_slots.items():
# #                 response += f"{day}: {', '.join(slots)}\n"
# #             return response.strip()
# #         else:
# #             if day_filter:
# #                 return f"✅ {faculty_name} has no free slots on {day_filter}."
# #             else:
# #                 return f"✅ {faculty_name} has no free slots."

# #     # Full timetable
# #     response = f"📅 Timetable for {faculty_name}:\n"
# #     for day, slots in timetable.items():
# #         response += f"{day}:\n"
# #         for hour, subject in slots.items():
# #             response += f"  {hour}: {subject}\n"
# #         response += "\n"
# #     return response.strip()
# import pandas as pd
# import re
# from datetime import datetime, time

# FILE_PATH = "data/workload.xlsx" #r"C:\Users\Administrator\Desktop\AU\Chatbot\workload.xlsx"

# HOURS = ["H1", "H2", "H3", "H4", "Lunch", "H5", "H6", "H7", "H8"]
# DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]

# _excel_cache = None


# # ------------------- EXCEL LOADER -------------------

# def load_excel():
#     global _excel_cache
#     if _excel_cache is None:
#         _excel_cache = pd.ExcelFile(FILE_PATH)
#     return _excel_cache


# # ------------------- BASIC HELPERS -------------------

# def find_best_faculty_match(query):
#     xls = load_excel()
#     query = query.lower()

#     for sheet in xls.sheet_names:
#         if sheet.lower() == query:
#             return sheet

#     for sheet in xls.sheet_names:
#         if sheet.lower() in query:
#             return sheet

#     query_words = query.split()
#     for sheet in xls.sheet_names:
#         if any(word in sheet.lower() for word in query_words):
#             return sheet

#     return None


# def get_faculty_timetable(faculty_name):
#     xls = load_excel()
#     sheet = find_best_faculty_match(faculty_name)

#     if not sheet:
#         return None, f"❌ Faculty '{faculty_name}' not found."

#     df = pd.read_excel(FILE_PATH, sheet_name=sheet, header=None)

#     timetable = {}
#     for i, day in enumerate(DAYS, start=1):
#         row_data = df.iloc[i, 1:10].tolist()
#         cleaned = ["Free" if (pd.isna(x) or str(x).strip() == "") else str(x) for x in row_data]
#         timetable[day] = dict(zip(HOURS, cleaned))

#     return timetable, sheet


# def get_free_slots(timetable, day_filter=None):
#     free = {}
#     for day, slots in timetable.items():
#         if day_filter and day.lower() != day_filter.lower():
#             continue

#         empty = [h for h, v in slots.items() if v == "Free" and h != "Lunch"]
#         if empty:
#             free[day] = empty
#     return free


# # ------------------- NEW FEATURE 1 -------------------
# #   CHECK WHO IS FREE RIGHT NOW
# # ----------------------------------------------------

# def get_current_hour():
#     """Map current time to timetable hour."""
#     now = datetime.now().time()

#     # You can modify these mappings to match your institution timings
#     hour_map = {
#         "H1": (time(9, 0), time(10, 0)),
#         "H2": (time(10, 0), time(11, 0)),
#         "H3": (time(11, 0), time(12, 0)),
#         "H4": (time(12, 0), time(13, 0)),
#         "Lunch": (time(13, 0), time(14, 0)),
#         "H5": (time(14, 0), time(15, 0)),
#         "H6": (time(15, 0), time(16, 0)),
#         "H7": (time(16, 0), time(17, 0)),
#         "H8": (time(17, 0), time(18, 0)),
#     }

#     for hour, (start, end) in hour_map.items():
#         if start <= now < end:
#             return hour
#     return None


# def who_is_free_now():
#     """Return faculty free at the current system time."""
#     day = datetime.now().strftime("%A")
#     if day not in DAYS:
#         return "Today is a holiday or weekend."

#     current_hour = get_current_hour()
#     if not current_hour:
#         return "Not within college hours."

#     xls = load_excel()
#     free_list = []

#     for sheet in xls.sheet_names:
#         timetable, _ = get_faculty_timetable(sheet)
#         if timetable and timetable.get(day, {}).get(current_hour) == "Free":
#             free_list.append(sheet)

#     if free_list:
#         return f"🕒 Free **now** ({day} {current_hour}): {', '.join(free_list)}"
#     return f"No faculty is free now ({day} {current_hour})."


# # ------------------- NEW FEATURE 2 -------------------
# #   WHO IS TAKING <SUBJECT> ON <DAY> <HOUR>
# # ----------------------------------------------------

# def who_is_teaching(subject, day, hour):
#     """Search all faculty for who teaches the specified subject."""
#     subject = subject.lower()
#     xls = load_excel()
#     found = []

#     for sheet in xls.sheet_names:
#         timetable, _ = get_faculty_timetable(sheet)
#         slot = timetable.get(day, {}).get(hour, "")

#         if subject in slot.lower():  # partial match
#             found.append(sheet)

#     return found


# # ------------------- QUERY PARSER + MAIN LOGIC -------------------

# def parse_workload_query(query):
#     q = query.lower()

#     faculty = find_best_faculty_match(q)
#     day = next((d for d in DAYS if d.lower() in q), None)
#     hour = next((h for h in HOURS if h.lower() in q), None)

#     # detect subject (word after "taking"/"teaching"/"class")
#     subject_match = re.search(r"(taking|teaching|class)\s+([a-zA-Z0-9 ]+)", q)
#     subject = subject_match.group(2).strip() if subject_match else None

#     return faculty, day, hour, subject


# def gui_get_workload(query):
#     faculty, day, hour, subject = parse_workload_query(query)
#     q = query.lower()

#     # ✨ 1. Check who is free NOW
#     if "free now" in q or "free right now" in q or "who is free now" in q:
#         return who_is_free_now()

#     # ✨ 2. Who is taking Python on Monday H2?
#     if subject and day and hour:
#         teachers = who_is_teaching(subject, day, hour)
#         if teachers:
#             return f"📘 {subject.title()} on {day} {hour} is taken by: {', '.join(teachers)}"
#         else:
#             return f"❌ No one is taking {subject.title()} on {day} {hour}."

#     # Existing features below...

#     # Case: Who is free on Monday H2?
#     if "who" in q and "free" in q and day and hour:
#         free_list = who_is_free(day, hour)
#         if free_list:
#             return f"🕒 Faculty free on {day} {hour}: {', '.join(free_list)}"
#         return f"❌ No one is free on {day} {hour}."

#     if not faculty:
#         return "❌ Please mention a faculty name or specify subject/day/hour."

#     timetable, sheet = get_faculty_timetable(faculty)
#     if timetable is None:
#         return sheet

#     if "free" in q:
#         free = get_free_slots(timetable, day)
#         if free:
#             msg = f"🕒 Free slots for {sheet}:\n"
#             for d, slots in free.items():
#                 msg += f"{d}: {', '.join(slots)}\n"
#             return msg.strip()
#         return f"❌ No free hours found for {sheet}."
#     # --- Day-specific timetable ---
#     if faculty and day and "free" not in q:
#         slots = timetable.get(day)
#         if slots:
#             msg = f"📅 Timetable for {sheet} on {day}:\n"
#             for h, s in slots.items():
#                 msg += f"  {h}: {s}\n"
#             return msg.strip()
#         else:
#             return f"❌ No timetable found for {sheet} on {day}."
#     # --- Free slots requested ---
#     if "free" in q:
        
#         # If a specific day is mentioned
#         if day:
#             day_slots = timetable.get(day)
#             if not day_slots:
#                 return f"❌ No data for {sheet} on {day}."

#             free_hours = [h for h, s in day_slots.items() if s == "Free" and h != "Lunch"]

#             if free_hours:
#                 return f"🕒 Free slots for {sheet} on {day}: {', '.join(free_hours)}"
#             else:
#                 return f"✅ {sheet} has no free slots on {day}."

#         # If user says "today"
#         if "today" in q:
#             today = datetime.now().strftime("%A")
#             if today not in DAYS:
#                 return "❌ Today is not a working day."

#             day_slots = timetable.get(today)
#             free_hours = [h for h, s in day_slots.items() if s == "Free" and h != "Lunch"]

#             if free_hours:
#                 return f"🕒 Free slots for {sheet} today ({today}): {', '.join(free_hours)}"
#             else:
#                 return f"✅ {sheet} has no free slots today."

#         # No specific day → return all days
#         free = get_free_slots(timetable)
#         if free:
#             msg = f"🕒 All free slots for {sheet}:\n"
#             for d, slots in free.items():
#                 msg += f"{d}: {', '.join(slots)}\n"
#             return msg.strip()
#         else:
#             return f"❌ {sheet} has no free slots on any day."
# def who_is_free(day, hour):
#     """Return faculty free on a specific day & hour."""
#         # ✨ NEW: "Who is free today?"
#     if "who" in q and "free" in q and "today" in q and not hour:
#         today = datetime.now().strftime("%A")
#         free_list = []
#         xls = load_excel()
#         current_timetable_hour = None

#         # Collect all who have at least one free slot today
#         for sheet in xls.sheet_names:
#             timetable, _ = get_faculty_timetable(sheet)
#             if timetable:
#                 slots = timetable.get(today, {})
#                 if any(v == "Free" for v in slots.values()):
#                     free_list.append(sheet)

#         if free_list:
#             return f"🕒 Faculty who have free hours today ({today}): {', '.join(free_list)}"
#         return f"❌ No one is free today ({today})."

#     xls = load_excel()
#     free_list = []

#     for sheet in xls.sheet_names:
#         timetable, _ = get_faculty_timetable(sheet)
#         if timetable and timetable.get(day, {}).get(hour) == "Free":
#             free_list.append(sheet)

#     return free_list
#     # Detect "today"
#     if "today" in q:
#         day = datetime.now().strftime("%A")

#     # Detect hour like H1, h2, hour 3
#     hour_match = re.search(r"h\s*([1-8])", q)
#     if hour_match:
#         hour = "H" + hour_match.group(1)

#     # Full timetable
#     msg = f"📅 Timetable for {sheet}:\n\n"
#     for d, slots in timetable.items():
#         msg += f"{d}:\n"
#         for h, s in slots.items():
#             msg += f"  {h}: {s}\n"
#         msg += "\n"
#     return msg.strip()

# import pandas as pd
# import re
# from datetime import datetime, time

# FILE_PATH = "data/workload.xlsx" 

# HOURS = ["H1", "H2", "H3", "H4", "Lunch", "H5", "H6", "H7", "H8"]
# DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]

# _excel_cache = None

# WORKLOAD_HELP_TEXT = """
# 📘 **Faculty Workload Help – Supported Queries**

# ==============================
# 🔹 Faculty Timetable
# - Harini timetable
# - Show timetable for Ramesh
# - Mamatha classes

# ==============================
# 🔹 Day Timetable
# - Harini Monday timetable
# - Ramesh Tuesday schedule

# ==============================
# 🔹 Who is Free?
# - Who is free today?
# - Who is free H2 today?
# - Who is free H3 on Friday?
# - Who is free H5 on Monday?

# ==============================
# 🔹 Free Slots of a Faculty
# - Harini free slots
# - Free slots for Mamatha on Tuesday
# - Ramesh free today

# ==============================
# 🔹 Who is Teaching a Subject?
# - Who is teaching Python on Monday H2?
# - Who is taking Data Science today H3?

# ==============================
# 🔹 Natural Language Queries
# - Is Harini free now?
# - Faculty free right now?
# =============================
# 🔹 Fuzzy Search Supported
# - Hareeni
# - Mamtha
# - Rames
# - H-2, hr2, hour2

# ==============================
# 🤖 Ask naturally!
# Example: "Who is free H3 today?"
# """

# # ------------------- EXCEL LOADER -------------------

# def load_excel():
#     global _excel_cache
#     if _excel_cache is None:
#         _excel_cache = pd.ExcelFile(FILE_PATH)
#     return _excel_cache


# # ------------------- BASIC HELPERS -------------------

# def find_best_faculty_match(query):
#     xls = load_excel()
#     query = query.lower()

#     # Exact
#     for sheet in xls.sheet_names:
#         if sheet.lower() == query:
#             return sheet

#     # Partial
#     for sheet in xls.sheet_names:
#         if sheet.lower() in query:
#             return sheet

#     # Split and match
#     query_words = query.split()
#     for sheet in xls.sheet_names:
#         if any(word in sheet.lower() for word in query_words):
#             return sheet

#     return None


# def get_faculty_timetable(faculty_name):
#     """
#     Auto-detect:
#     - Day row index (row containing Monday, Tuesday, ...)
#     - Period columns (columns after the first non-empty day column)
#     - Missing cells → Free
#     """

#     xls = load_excel()
#     sheet = find_best_faculty_match(faculty_name)

#     if not sheet:
#         return None, f"❌ Faculty '{faculty_name}' not found."

#     df = pd.read_excel(FILE_PATH, sheet_name=sheet, header=None)

#     # Normalize all cells
#     df_filled = df.fillna("").astype(str)

#     # ----- 1️⃣ Identify the DAY column and DAY row indexes -----
#     day_row_map = {}   # day → row index
#     day_column = None  # where Monday/Tuesdays appear

#     for row_idx in range(len(df_filled)):
#         for col_idx in range(len(df_filled.columns)):
#             cell = df_filled.iat[row_idx, col_idx].strip().lower()

#             if cell in [d.lower() for d in DAYS]:
#                 day_name = cell.capitalize()
#                 day_row_map[day_name] = row_idx

#                 if day_column is None:
#                     day_column = col_idx  # first column where a day is written

#     if not day_row_map:
#         return None, f"❌ Could not detect day rows in '{sheet}'. Excel format issue."

#     # ----- 2️⃣ Detect period columns -----
#     # Period columns = all columns AFTER the day column
#     period_columns = list(range(day_column + 1, len(df_filled.columns)))
#     num_periods = len(period_columns)

#     if num_periods < 1:
#         return None, f"❌ No period columns found for '{sheet}'."

#     # Trim or extend HOURS to match detected columns
#     dynamic_hours = HOURS[:num_periods]

#     # ----- 3️⃣ Build CLEAN TIMETABLE -----
#     timetable = {}

#     for day, row_idx in day_row_map.items():
#         day_periods = {}

#         for hour_label, col_idx in zip(dynamic_hours, period_columns):

#             # safe cell fetch
#             try:
#                 val = df_filled.iat[row_idx, col_idx].strip()
#                 val = "Free" if val == "" else val
#             except:
#                 val = "Free"

#             day_periods[hour_label] = val

#         timetable[day] = day_periods

#     return timetable, sheet



# # ------------------- FREE SLOT UTILITIES -------------------

# def get_free_slots(timetable, day_filter=None):
#     free = {}
#     for day, slots in timetable.items():
#         if day_filter and day.lower() != day_filter.lower():
#             continue

#         empty = [h for h, v in slots.items() if v == "Free" and h != "Lunch"]
#         if empty:
#             free[day] = empty
#     return free


# # ------------------- NEW: FREE NOW FINDER -------------------

# def get_current_hour():
#     now = datetime.now().time()

#     hour_map = {
#         "H1": (time(9, 0), time(10, 0)),
#         "H2": (time(10, 0), time(11, 0)),
#         "H3": (time(11, 0), time(12, 0)),
#         "H4": (time(12, 0), time(13, 0)),
#         "Lunch": (time(13, 0), time(14, 0)),
#         "H5": (time(14, 0), time(15, 0)),
#         "H6": (time(15, 0), time(16, 0)),
#         "H7": (time(16, 0), time(17, 0)),
#         "H8": (time(17, 0), time(18, 0)),
#     }

#     for hour, (start, end) in hour_map.items():
#         if start <= now < end:
#             return hour
#     return None


# def who_is_free_now():
#     day = datetime.now().strftime("%A")
#     if day not in DAYS:
#         return "Today is a holiday or weekend."

#     current_hour = get_current_hour()
#     if not current_hour:
#         return "Not within college hours."

#     xls = load_excel()
#     free_list = []

#     for sheet in xls.sheet_names:
#         timetable, _ = get_faculty_timetable(sheet)
#         if timetable and timetable.get(day, {}).get(current_hour) == "Free":
#             free_list.append(sheet)

#     if free_list:
#         return f"🕒 Free **now** ({day} {current_hour}): {', '.join(free_list)}"
#     return f"No faculty is free now ({day} {current_hour})."


# # ------------------ NEW: FREE ON SPECIFIC DAY & HOUR ------------------

# def who_is_free(day, hour):
#     xls = load_excel()
#     free_list = []

#     for sheet in xls.sheet_names:
#         timetable, _ = get_faculty_timetable(sheet)
#         if timetable and timetable.get(day, {}).get(hour) == "Free":
#             free_list.append(sheet)

#     return free_list


# # ------------------- SUBJECT / TEACHER MAPPING -------------------

# def who_is_teaching(subject, day, hour):
#     subject = subject.lower()
#     xls = load_excel()
#     found = []

#     for sheet in xls.sheet_names:
#         timetable, _ = get_faculty_timetable(sheet)
#         slot = timetable.get(day, {}).get(hour, "")

#         if subject in slot.lower():
#             found.append(sheet)

#     return found


# # ------------------- QUERY PARSER -------------------
# def is_workload_help(query: str):
#     q = query.lower()
#     help_keywords = ["help", "commands", "usage", "how to", "what can", "workload help", "faculty help"]
#     return any(k in q for k in help_keywords)

# def parse_workload_query(query):
#     q = query.lower()

#     faculty = find_best_faculty_match(q)
#     day = next((d for d in DAYS if d.lower() in q), None)
#     hour = next((h for h in HOURS if h.lower() in q), None)

#     # Detect "today"
#     if "today" in q:
#         day = datetime.now().strftime("%A")

#     # Detect H1/H2/H3 etc
#     hour_match = re.search(r"h\s*([1-8])", q)
#     if hour_match:
#         hour = "H" + hour_match.group(1)

#     # Detect subject
#     subject_match = re.search(r"(taking|teaching|class)\s+([a-zA-Z0-9 ]+)", q)
#     subject = subject_match.group(2).strip() if subject_match else None

#     return faculty, day, hour, subject


# # ------------------- MAIN ENGINE -------------------
   

# def gui_get_workload(query):
#         # ===============================
#     # NEW: "Who has no class in H4?"
#     # ===============================
#     if "who" in q and ("no class" in q or "doesnt have class" in q or "don't have class" in q or "free" in q) and hour and day is None:
#         # default day = today
#         today = datetime.now().strftime("%A")
#         free_list = who_is_free(today, hour)

#         if free_list:
#             return f"🕒 Faculty with NO class in {hour} today ({today}): {', '.join(free_list)}"
#         return f"❌ No one is free in {hour} today ({today})."


#     # ===========================================
#     # NEW: "Who is teaching python H2?"
#     # ===========================================
#     if ("who" in q and ("teaching" in q or "taking" in q or "handles" in q)) and subject and hour and day is None:
#         # default day = today
#         today = datetime.now().strftime("%A")
#         teachers = who_is_teaching(subject, today, hour)

#         if teachers:
#             return f"📘 {subject.title()} in {hour} today ({today}) is taken by: {', '.join(teachers)}"
#         else:
#             return f"❌ No one is teaching {subject.title()} in {hour} today ({today})."

#      # 🆘 HELP MENU
#     if is_workload_help(query):
#         return WORKLOAD_HELP_TEXT
#     faculty, day, hour, subject = parse_workload_query(query)
#     q = query.lower()

#     # 1️⃣ Who is free NOW?
#     if "free now" in q or "free right now" in q or "who is free now" in q:
#         return who_is_free_now()

#     # 2️⃣ Who is teaching <subject> on Monday H2?
#     if subject and day and hour:
#         teachers = who_is_teaching(subject, day, hour)
#         if teachers:
#             return f"📘 {subject.title()} on {day} {hour} is taken by: {', '.join(teachers)}"
#         else:
#             return f"❌ No one is taking {subject.title()} on {day} {hour}."

#     # 3️⃣ Who is free today?  (no hour)
#     if "who" in q and "free" in q and "today" in q and not hour:
#         today = datetime.now().strftime("%A")
#         xls = load_excel()
#         free_list = []

#         for sheet in xls.sheet_names:
#             timetable, _ = get_faculty_timetable(sheet)
#             slots = timetable.get(today, {})
#             if any(v == "Free" for v in slots.values()):
#                 free_list.append(sheet)

#         if free_list:
#             return f"🕒 Faculty who have free hours today ({today}): {', '.join(free_list)}"
#         return f"❌ No one is free today ({today})."

#     # 4️⃣ Who is free H2 today?
#     if "who" in q and "free" in q and hour and "today" in q:
#         today = datetime.now().strftime("%A")
#         free_list = who_is_free(today, hour)

#         if free_list:
#             return f"🕒 Faculty free in {hour} today ({today}): {', '.join(free_list)}"
#         return f"❌ No one is free in {hour} today."

#     # 5️⃣ Who is free H2 on Monday?
#     if "who" in q and "free" in q and day and hour and "today" not in q:
#         free_list = who_is_free(day, hour)
#         if free_list:
#             return f"🕒 Faculty free on {day} at {hour}: {', '.join(free_list)}"
#         return f"❌ No one is free on {day} at {hour}."

#     # Existing behavior...

#     if not faculty:
#         return "❌ Please mention a faculty name or specify subject/day/hour."

#     timetable, sheet = get_faculty_timetable(faculty)
#     if timetable is None:
#         return sheet

#     # Free slots for a day
#     if "free" in q:
#         if day:
#             free_hours = [h for h, s in timetable.get(day, {}).items() if s == "Free" and h != "Lunch"]
#             if free_hours:
#                 return f"🕒 Free slots for {sheet} on {day}: {', '.join(free_hours)}"
#             else:
#                 return f"❌ {sheet} has no free slots on {day}."

#         # Today free slots
#         if "today" in q:
#             today = datetime.now().strftime("%A")
#             free_hours = [h for h, s in timetable.get(today, {}).items() if s == "Free" and h != "Lunch"]
#             if free_hours:
#                 return f"🕒 Free slots for {sheet} today ({today}): {', '.join(free_hours)}"
#             else:
#                 return f"❌ {sheet} has no free slots today."

#         free = get_free_slots(timetable)
#         if free:
#             msg = f"🕒 All free slots for {sheet}:\n"
#             for d, slots in free.items():
#                 msg += f"{d}: {', '.join(slots)}\n"
#             return msg.strip()
#         else:
#             return f"❌ {sheet} has no free slots on any day."

#     # Day-specific timetable
#     if faculty and day and "free" not in q:
#         slots = timetable.get(day)
#         if slots:
#             msg = f"📅 Timetable for {sheet} on {day}:\n"
#             for h, s in slots.items():
#                 msg += f"  {h}: {s}\n"
#             return msg.strip()

#     # Full timetable
#     msg = f"📅 Timetable for {sheet}:\n\n"
#     for d, slots in timetable.items():
#         msg += f"{d}:\n"
#         for h, s in slots.items():
#             msg += f"  {h}: {s}\n"
#         msg += "\n"
#     return msg.strip()
import pandas as pd
import re
from datetime import datetime, time

# ✔ Correct path
FILE_PATH = "data/workload.xlsx"

HOURS = ["H1", "H2", "H3", "H4", "Lunch", "H5", "H6", "H7", "H8"]
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]

_excel_cache = None


# =====================================================================
#                            HELP TEXT
# =====================================================================

WORKLOAD_HELP_TEXT = """
📘 **Faculty Workload Help – Supported Queries**

==============================
🔹 Faculty Timetable
- Harini timetable
- Show timetable for Ramesh
- Mamatha classes

==============================
🔹 Day Timetable
- Harini Monday timetable
- Ramesh Tuesday schedule

==============================
🔹 Who is Free?
- Who is free today?
- Who is free H2 today?
- Who is free H3 on Friday?
- Who is free H5 on Monday?
- Who is free in H3?
==============================
🔹 Free Slots (Individual Faculty)
- Harini free slots
- Free slots for Mamatha on Tuesday
- Ramesh free today

==============================

🔹 Natural Language Queries
- Is Harini free now?
- Faculty free right now?
==============================
🔹 Fuzzy Search Supported
- Hareeni
- Mamtha
- Rames
- H2, hr2, hour2

==============================
🤖 Ask naturally!
Example: "Who is free H3 today?"
"""


def is_workload_help(query: str):
    q = query.lower()
    help_keywords = ["help", "commands", "usage", "how to", "workload help", "faculty help"]
    return any(k in q for k in help_keywords)


# =====================================================================
#                      EXCEL LOADING + NORMALIZATION
# =====================================================================

def load_excel():
    global _excel_cache
    if _excel_cache is None:
        _excel_cache = pd.ExcelFile(FILE_PATH)
    return _excel_cache


# =====================================================================
#                      FACULTY NAME MATCHING
# =====================================================================

def find_best_faculty_match(query):
    xls = load_excel()
    q = query.lower()

    # Exact match
    for sheet in xls.sheet_names:
        if sheet.lower() == q:
            return sheet

    # Partial in query
    for sheet in xls.sheet_names:
        if sheet.lower() in q:
            return sheet

    # Match any word
    q_words = q.split()
    for sheet in xls.sheet_names:
        if any(word in sheet.lower() for word in q_words):
            return sheet

    return None


# =====================================================================
#                      AUTO-DETECT TIMETABLE STRUCTURE
# =====================================================================

def get_faculty_timetable(faculty_name):
    """
    Auto-detect:
    ✔ Day rows
    ✔ Period columns
    ✔ Missing values handled
    """

    xls = load_excel()
    sheet = find_best_faculty_match(faculty_name)

    if not sheet:
        return None, f"❌ Faculty '{faculty_name}' not found."

    df = pd.read_excel(FILE_PATH, sheet_name=sheet, header=None)
    df = df.fillna("").astype(str)

    # -------- 1️⃣ Detect day rows + day column --------
    day_row_map = {}
    day_col = None

    for r in range(len(df)):
        for c in range(len(df.columns)):
            cell = df.iat[r, c].strip().lower()
            if cell in [d.lower() for d in DAYS]:
                day_row_map[cell.capitalize()] = r
                if day_col is None:
                    day_col = c

    if not day_row_map:
        return None, f"❌ Could not detect day rows for {sheet}. Check Excel format."

    # -------- 2️⃣ Detect period columns (right of day column) --------
    period_cols = list(range(day_col + 1, len(df.columns)))
    dynamic_hours = HOURS[:len(period_cols)]

    # -------- 3️⃣ Build timetable dictionary --------
    timetable = {}

    for day, row in day_row_map.items():
        day_slots = {}

        for hour, col in zip(dynamic_hours, period_cols):
            try:
                value = df.iat[row, col].strip()
                value = value if value else "Free"
            except:
                value = "Free"

            day_slots[hour] = value

        timetable[day] = day_slots

    return timetable, sheet


# =====================================================================
#                   FREE SLOT UTILITIES + CURRENT HOUR
# =====================================================================

def get_free_slots(timetable, day_filter=None):
    free = {}
    for day, slots in timetable.items():
        if day_filter and day.lower() != day_filter.lower():
            continue
        empty = [h for h, v in slots.items() if v == "Free" and h != "Lunch"]
        if empty:
            free[day] = empty
    return free


def get_current_hour():
    now = datetime.now().time()
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
    for h, (start, end) in hour_map.items():
        if start <= now < end:
            return h
    return None


def who_is_free(day, hour):
    xls = load_excel()
    free_list = []
    for sheet in xls.sheet_names:
        timetable, _ = get_faculty_timetable(sheet)
        if timetable is None:
            continue
        if timetable.get(day, {}).get(hour) == "Free":
            free_list.append(sheet)
    return free_list


def who_is_free_now():
    day = datetime.now().strftime("%A")
    hour = get_current_hour()

    if not hour:
        return "Not within college hours."

    free_list = who_is_free(day, hour)
    if free_list:
        return f"🕒 Free NOW ({day} {hour}): {', '.join(free_list)}"
    return f"❌ No faculty free now ({day} {hour})"


# =====================================================================
#                   WHO IS TEACHING A SUBJECT?
# =====================================================================

def who_is_teaching(subject, day, hour):
    subject = subject.lower()
    xls = load_excel()
    found = []

    for sheet in xls.sheet_names:
        timetable, _ = get_faculty_timetable(sheet)
        if timetable is None:
            continue

        slot = timetable.get(day, {}).get(hour, "")
        if subject in slot.lower():
            found.append(sheet)

    return found


# =====================================================================
#                   QUERY PARSER
# =====================================================================

def parse_workload_query(query):
    q = query.lower()

    faculty = find_best_faculty_match(q)
    day = next((d for d in DAYS if d.lower() in q), None)
    hour = next((h for h in HOURS if h.lower() in q), None)

    # Detect H2, h2, hour2
    m = re.search(r"h\s*([1-8])", q)
    if m:
        hour = "H" + m.group(1)

    # Detect "today"
    if "today" in q:
        day = datetime.now().strftime("%A")

    # Detect subject after "teaching", "taking"
    m2 = re.search(r"(teaching|taking|class|handles)\s+([a-zA-Z0-9 ]+)", q)
    subject = m2.group(2).strip() if m2 else None

    return faculty, day, hour, subject


# =====================================================================
#                   MAIN WORKLOAD ENGINE
# =====================================================================

def gui_get_workload(query):
    q = query.lower()   # IMPORTANT: define first
    faculty, day, hour, subject = parse_workload_query(query)

    # 🆘 Help command
    if is_workload_help(query):
        return WORKLOAD_HELP_TEXT

    # 1️⃣ Who is free NOW
    if "free now" in q or "free right now" in q or "who is free now" in q:
        return who_is_free_now()

    # 2️⃣ Who is teaching Python Monday H2 (full)
    if subject and day and hour:
        teachers = who_is_teaching(subject, day, hour)
        if teachers:
            return f"📘 {subject.title()} on {day} {hour} is taken by: {', '.join(teachers)}"
        return f"❌ No one teaches {subject.title()} on {day} {hour}"

    # 3️⃣ Who is free today (no hour)
    if "who" in q and "free" in q and "today" in q and not hour:
        today = datetime.now().strftime("%A")
        free_today = []
        xls = load_excel()

        for sheet in xls.sheet_names:
            timetable, _ = get_faculty_timetable(sheet)
            if timetable is None:
                continue
            if any(v == "Free" for v in timetable.get(today, {}).values()):
                free_today.append(sheet)

        if free_today:
            return f"🕒 Faculty with free hour today ({today}): {', '.join(free_today)}"
        return f"❌ No one is free today."

    # 4️⃣ Who is free H2 today?
    if "who" in q and "free" in q and hour and "today" in q:
        today = datetime.now().strftime("%A")
        free_list = who_is_free(today, hour)

        if free_list:
            return f"🕒 Free in {hour} today ({today}): {', '.join(free_list)}"
        return f"❌ No one is free in {hour} today."

    # 5️⃣ Who is free H2 on Monday?
    if "who" in q and "free" in q and day and hour and "today" not in q:
        free_list = who_is_free(day, hour)
        if free_list:
            return f"🕒 Faculty free on {day} at {hour}: {', '.join(free_list)}"
        return f"❌ No one is free on {day} at {hour}."

    # 6️⃣ NEW: "Who has no class in H4?"
    if "who" in q and ("no class" in q or "doesnt have class" in q or "don't have class" in q) and hour:
        today = datetime.now().strftime("%A")
        free_list = who_is_free(today, hour)

        if free_list:
            return f"🕒 Faculty with NO class in {hour} today ({today}): {', '.join(free_list)}"
        return f"❌ Everyone has class in {hour} today."

    # 7️⃣ NEW: "Who is teaching python H2?"
    if ("who" in q and ("teaching" in q or "taking" in q or "handles" in q)) and subject and hour and day is None:
        today = datetime.now().strftime("%A")
        teachers = who_is_teaching(subject, today, hour)

        if teachers:
            return f"📘 {subject.title()} in {hour} today ({today}) is taken by: {', '.join(teachers)}"
        return f"❌ No one is teaching {subject.title()} in {hour} today."

    # Existing logic...

    if not faculty:
        return "❌ Please specify a faculty name or a day/hour query."

    timetable, sheet = get_faculty_timetable(faculty)
    if timetable is None:
        return sheet

    # Faculty free slots (all days)
    if "free" in q and not day:
        free = get_free_slots(timetable)
        if free:
            msg = f"🕒 Free slots for {sheet}:\n"
            for d, slots in free.items():
                msg += f"{d}: {', '.join(slots)}\n"
            return msg
        return f"❌ {sheet} has no free hours."

    # Free slots for specific day
    if "free" in q and day:
        slots = timetable.get(day, {})
        free_hours = [h for h, v in slots.items() if v == "Free" and h != "Lunch"]

        if free_hours:
            return f"🕒 Free slots for {sheet} on {day}: {', '.join(free_hours)}"
        return f"❌ {sheet} has no free slots on {day}."

    # Day timetable
    if faculty and day:
        slots = timetable.get(day)
        if not slots:
            return f"❌ No timetable for {sheet} on {day}."
        msg = f"📅 Timetable for {sheet} on {day}:\n"
        for h, s in slots.items():
            msg += f"  {h}: {s}\n"
        return msg

    # Full timetable
    msg = f"📅 Timetable for {sheet}:\n\n"
    for d, slot in timetable.items():
        msg += f"{d}:\n"
        for h, s in slot.items():
            msg += f"  {h}: {s}\n"
        msg += "\n"
    return msg  