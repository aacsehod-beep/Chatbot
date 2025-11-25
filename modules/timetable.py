# import pandas as pd

# # ----------------- Load Excel -----------------
# file_path = r"C:\Users\Administrator\Desktop\AU\Chatbot\timetable.xlsx"
# xls = pd.ExcelFile(file_path)

# # ----------------- Parse all sheets -----------------
# timetables = {}

# for sheet in xls.sheet_names:
#     df = pd.read_excel(file_path, sheet_name=sheet, header=None)
#     section = sheet.strip().lower()

#     # Days in column A, rows 4-8 (0-indexed 3:8)
#     days = df.iloc[3:8, 0].tolist()

#     # Hours in row 3, columns B-J (B=1 ... J=9)
#     hours_labels = ['H1','H2','H3','H4','H0','H5','H6','H7','H8']  # H0 = Lunch Break
#     hours_cols = [1,2,3,4,5,6,7,8,9]

#     timetable = {}
#     for i, day in enumerate(days):
#         timetable[day] = {}
#         for col_idx, hour_label in zip(hours_cols, hours_labels):
#             timetable[day][hour_label] = df.iloc[i+3, col_idx]  # classes start from row 4

#     # Class Incharge
#     class_incharge = None
#     for row in df.iloc[:, 0]:
#         if isinstance(row, str) and "class incharge" in row.lower():
#             class_incharge = row.split(":")[-1].strip()

#     timetables[section] = {"class_incharge": class_incharge, "schedule": timetable}

# # ----------------- GUI-friendly query -----------------
# def gui_get_timetable(query):
#     """
#     Query examples:
#     - "CSE-2A timetable" (full week)
#     - "CSE-2A Monday timetable" (specific day)
#     - "CSE-2A H4 Thursday" (specific hour)
#     - "CSE-2A free slots Friday"
#     - "CSE-2A class incharge"
#     """
#     query_lower = query.lower()
#     section = None

#     # Identify section
#     for sec in timetables.keys():
#         if sec in query_lower.replace(" ", ""):
#             section = sec
#             break
#     if not section:
#         return ["⚠️ Please specify a valid section (e.g., CSE-2A)."]

#     data = timetables[section]
#     schedule = data["schedule"]
#     response_lines = []

#     # Class Incharge
#     if "class incharge" in query_lower:
#         response_lines.append(f"👨‍🏫 Class Incharge of {section.upper()}: {data['class_incharge']}")
#         return response_lines

#     # Full timetable
#     if "timetable" in query_lower and not any(day.lower() in query_lower for day in schedule.keys()):
#         response_lines.append(f"📅 Timetable for {section.upper()}:")
#         for day, hours in schedule.items():
#             response_lines.append(f"\n--- {day} ---")
#             for h, sub in hours.items():
#                 response_lines.append(f"{h}: {sub}")
#         return response_lines

#     # Day-specific timetable
#     for day in schedule.keys():
#         if day.lower() in query_lower and "timetable" in query_lower:
#             response_lines.append(f"📅 {section.upper()} - {day} Timetable:")
#             for h, sub in schedule[day].items():
#                 response_lines.append(f"{h}: {sub}")
#             return response_lines

#     # Hour-specific
#     for day in schedule.keys():
#         if day.lower() in query_lower:
#             for h, sub in schedule[day].items():
#                 if h.lower() in query_lower:
#                     return [f"{section.upper()} - {day} {h}: {sub}"]

#     # Free slots
#     if "free slot" in query_lower or "free slots" in query_lower:
#         for day, hours in schedule.items():
#             if day.lower() in query_lower:
#                 free = [h for h, sub in hours.items() if isinstance(sub, str) and "free" in sub.lower()]
#                 return [f"🕒 Free slots in {section.upper()} on {day}: {', '.join(free) if free else 'None'}"]

#     # Subject-based search
#     for day, hours in schedule.items():
#         if day.lower() in query_lower:
#             for h, sub in hours.items():
#                 if isinstance(sub, str) and any(word in query_lower for word in sub.lower().split()):
#                     return [f"{section.upper()} has {sub} on {day} at {h}"]

#     return ["🤖 Sorry, I couldn't understand your question."]

# # ----------------- Example Usage -----------------
# if __name__ == "__main__":
#     queries = [
#         "CSE-2A timetable",
#         "CSE-2A Monday timetable",
#         "CSE-2A H4 Thursday",
#         "CSE-2A free slots Friday",
#         "CSE-2A class incharge"
#     ]
#     for q in queries:
#         lines = gui_get_timetable(q)
#         print("\n".join(lines))
#         print("\n" + "-"*50 + "\n")
# """
# timetable.py — Hybrid NLP-style Timetable Engine (A+B+C)
# Features:
# - Exact + fuzzy matching for sections, days, hours, teachers, subjects
# - Class incharge extraction from A10
# - Teacher-based search, subject search, free-slot finder
# - Next-class finder (uses Asia/Kolkata timezone)
# - Works with your existing Excel layout (A4:A8 days, B-J hours, A10 class incharge)
# """

# import pandas as pd
# import re
# import difflib
# from datetime import datetime, time
# try:
#     from zoneinfo import ZoneInfo  # Python 3.9+
# except Exception:
#     ZoneInfo = None

# # ---------------- CONFIG: map hour labels to approximate start times ----------------
# HOUR_STARTS = {
#     "H1": time(9, 0),
#     "H2": time(9, 50),
#     "H3": time(10, 40),
#     "H4": time(11, 30),
#     "H0": time(12, 20),  # Lunch
#     "H5": time(13, 10),
#     "H6": time(14, 0),
#     "H7": time(14, 50),
#     "H8": time(15, 40),
# }

# # ---------------- Utility helpers ----------------
# def normalize_text(s: str) -> str:
#     if s is None:
#         return ""
#     s = str(s).lower()
#     s = re.sub(r'[\u200b-\u200d\ufeff]', '', s)  # remove stray invisible chars
#     s = re.sub(r'[\s\-_\./,]+', ' ', s)  # normalize separators
#     s = re.sub(r'[^a-z0-9 ]', '', s)  # remove punctuation (keep alphanum + spaces)
#     s = s.strip()
#     return s

# def simple_tokens(s: str):
#     return [t for t in normalize_text(s).split() if t]

# def fuzzy_score(a: str, b: str) -> float:
#     """Return fuzzy similarity ratio (0..1) using difflib"""
#     if not a or not b:
#         return 0.0
#     return difflib.SequenceMatcher(None, a, b).ratio()

# def fuzzy_in(text: str, candidate: str, threshold=0.75) -> bool:
#     """Check if candidate is approximately inside text"""
#     text_n = normalize_text(text)
#     candidate_n = normalize_text(candidate)
#     if candidate_n in text_n:
#         return True
#     # check token-level match
#     if fuzzy_score(text_n, candidate_n) >= threshold:
#         return True
#     # check sliding substrings
#     words = text_n.split()
#     if len(candidate_n.split()) <= 1:
#         # compare individually
#         return any(fuzzy_score(w, candidate_n) >= threshold for w in words)
#     # multi-word: check n-gram windows
#     m = len(candidate_n.split())
#     for i in range(len(words) - m + 1):
#         window = " ".join(words[i:i+m])
#         if fuzzy_score(window, candidate_n) >= threshold:
#             return True
#     return False

# # ----------------- Load Excel -----------------
# FILE_PATH = r"C:\Users\Administrator\Desktop\AU\Chatbot\timetable.xlsx"

# xls = pd.ExcelFile(FILE_PATH)

# # main structure
# timetables = {}  # section -> {"class_incharge": str, "schedule": {day: {H1:cell,...}}}
# # indices for fast fuzzy search
# _subject_index = {}   # subject_lower -> list of (section, day, hour_label, raw_cell)
# _teacher_index = {}   # teacher_token -> list of (section, day, hour_label, raw_cell)
# _all_cells = []       # list of (section, day, hour_label, raw_cell)

# # Normalize day names for detection
# DAY_NAMES = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
# DAY_ABBRS = {d[:3]: d for d in DAY_NAMES}

# # hours mapping (columns B-J => indices 1..9)
# HOUR_LABELS = ['H1','H2','H3','H4','H0','H5','H6','H7','H8']
# HOUR_COLS = [1,2,3,4,5,6,7,8,9]

# # ----------------- Parse all sheets and build indices -----------------
# for sheet in xls.sheet_names:
#     df = pd.read_excel(FILE_PATH, sheet_name=sheet, header=None)
#     section = normalize_text(sheet).replace(" ", "")  # canonical section id e.g., cse2a

#     # read days A4:A8 (0-indexed rows 3..7)
#     try:
#         days = [str(x).strip() for x in df.iloc[3:8, 0].tolist()]
#     except Exception:
#         days = []

#     timetable = {}
#     for i, day in enumerate(days):
#         # keep original day name as in sheet
#         day_name = day if isinstance(day, str) and day.strip() else DAY_NAMES[i]  # fallback
#         timetable[day_name] = {}
#         for col_idx, hour_label in zip(HOUR_COLS, HOUR_LABELS):
#             try:
#                 raw_cell = df.iloc[i+3, col_idx]
#             except Exception:
#                 raw_cell = None
#             timetable[day_name][hour_label] = raw_cell
#             _all_cells.append((section, day_name, hour_label, raw_cell))

#             # index subjects and teachers: naive heuristic - any text that has letters
#             if isinstance(raw_cell, str) and raw_cell.strip():
#                 cell_norm = normalize_text(raw_cell)

#                 # build subject index: token sequences longer than 2 chars
#                 tokens = cell_norm.split()
#                 # heuristic: treat the part before comma as subject/descr
#                 primary_part = re.split(r',', raw_cell)[0]
#                 subkey = normalize_text(primary_part)
#                 if subkey:
#                     _subject_index.setdefault(subkey, []).append((section, day_name, hour_label, raw_cell))

#                 # also index n-grams for subjects up to length 4 tokens
#                 max_ng = min(4, len(tokens))
#                 for n in range(1, max_ng+1):
#                     for start in range(0, len(tokens)-n+1):
#                         gram = " ".join(tokens[start:start+n])
#                         if len(gram) >= 2:
#                             _subject_index.setdefault(gram, []).append((section, day_name, hour_label, raw_cell))

#                 # teacher heuristics: detect tokens that look like names (capitalization in raw, presence of titles)
#                 # We'll also index all multi-word sequences as possible teacher phrases
#                 raw_tokens = [t.strip() for t in re.split(r'[,/();\-]', raw_cell) if t.strip()]
#                 for chunk in raw_tokens:
#                     chunk_n = normalize_text(chunk)
#                     if len(chunk_n) > 2:
#                         _teacher_index.setdefault(chunk_n, []).append((section, day_name, hour_label, raw_cell))

#     # class incharge from A10 (row index 9, col 0)
#     class_incharge = None
#     try:
#         cell_value = df.iloc[9, 0]  # A10
#         if isinstance(cell_value, str) and "class incharge" in cell_value.lower():
#             class_incharge = cell_value.split(":", 1)[-1].strip()
#         else:
#             # fallback: if A10 is a name directly
#             if isinstance(cell_value, str) and cell_value.strip():
#                 class_incharge = cell_value.strip()
#     except Exception:
#         class_incharge = None

#     timetables[section] = {
#         "class_incharge": class_incharge,
#         "schedule": timetable
#     }

# # ----------------- Helper functions for query understanding -----------------
# def find_section_in_query(query: str):
#     q = normalize_text(query).replace(" ", "")
#     # direct match
#     if q in timetables:
#         return q
#     # fuzzy match against keys
#     keys = list(timetables.keys())
#     best = difflib.get_close_matches(q, keys, n=1, cutoff=0.6)
#     if best:
#         return best[0]
#     # try to find a key that is substring of q or vice versa
#     for k in keys:
#         if k in q or q in k:
#             return k
#     return None

# def find_day_in_query(query: str):
#     q = normalize_text(query)
#     # direct day names or abbr
#     for d in DAY_NAMES:
#         if d in q:
#             return d
#     for ab, full in DAY_ABBRS.items():
#         if ab in q:
#             return full
#     # fuzzy day matching
#     for d in DAY_NAMES:
#         if fuzzy_score(q, d) > 0.6 or fuzzy_score(q, d[:3]) > 0.8:
#             return d
#     return None

# def find_hour_label_in_query(query: str):
#     q = normalize_text(query)
#     # common patterns: h1, h2, hour 1, first hour
#     m = re.search(r'\b(h|hour|hr)?\s*([1-8])\b', query.lower())
#     if m:
#         num = int(m.group(2))
#         label = f"H{num}" if num != 0 else "H0"
#         if label in HOUR_LABELS:
#             return label
#     # ordinal words
#     ord_map = {
#         "first": "H1","second":"H2","third":"H3","fourth":"H4",
#         "fifth":"H5","sixth":"H6","seventh":"H7","eighth":"H8"
#     }
#     for word, lab in ord_map.items():
#         if word in q:
#             return lab
#     # explicit H0 lunch
#     if "lunch" in q or "break" in q:
#         return "H0"
#     # fuzzy Hx like h4, h5
#     for lab in HOUR_LABELS:
#         if lab.lower() in q:
#             return lab
#     return None

# def detect_incharge_request(query: str):
#     q = normalize_text(query)
#     keywords = ["incharge","in charge","class teacher","class incharge","class teacher","coordinator","mentor","class teacher","class incharge","class coordinator","class teacher","class mentor"]
#     return any(k.replace(" ", "") in q.replace(" ", "") or k in q for k in keywords)

# def detect_free_request(query: str):
#     q = normalize_text(query)
#     return any(w in q for w in ["free","free slot","free slots","break","available","vacant","no class","off time"])

# def detect_next_request(query: str):
#     q = normalize_text(query)
#     return any(w in q for w in ["next","upcoming","now","current","what is next","what's next","next class"])

# def find_subject_slots(query: str, section=None, day=None, top_n=10):
#     q = normalize_text(query)
#     results = []
#     # check subject index keys with fuzzy
#     for key, slots in _subject_index.items():
#         if fuzzy_score(q, key) >= 0.6 or key in q or q in key:
#             for s in slots:
#                 sec, d, h, raw = s
#                 if section and sec != section:
#                     continue
#                 if day and d.lower() != day.lower():
#                     continue
#                 results.append((sec, d, h, raw, fuzzy_score(q, key)))
#     # sort by score
#     results = sorted(results, key=lambda x: x[4], reverse=True)
#     return results[:top_n]

# def find_teacher_slots(query: str, section=None, day=None, top_n=20):
#     q = normalize_text(query)
#     candidates = []
#     # direct teacher index matching
#     for key, slots in _teacher_index.items():
#         score = fuzzy_score(q, key)
#         if score >= 0.55 or key in q:
#             for s in slots:
#                 sec, d, h, raw = s
#                 if section and sec != section:
#                     continue
#                 if day and d.lower() != day.lower():
#                     continue
#                 candidates.append((sec, d, h, raw, score))
#     # also check all cells by fuzzy_in (for cases where teacher phrase is buried)
#     if not candidates:
#         for sec, d, h, raw in _all_cells:
#             if not raw or not isinstance(raw, str):
#                 continue
#             raw_n = normalize_text(raw)
#             if fuzzy_in(raw_n, q, threshold=0.6):
#                 candidates.append((sec, d, h, raw, fuzzy_score(raw_n, q)))
#     candidates = sorted(candidates, key=lambda x: x[4], reverse=True)
#     return candidates[:top_n]

# # ----------------- Main chatbot-facing function -----------------
# def gui_get_timetable(query: str):
#     """
#     The generalized query processor. Accepts many natural styles.
#     Returns a list of response lines.
#     """
#     if not query or not query.strip():
#         return ["⚠️ Please type a question about timetable (e.g., 'CSE-2A Monday timetable')."]

#     raw_q = query.strip()
#     qnorm = normalize_text(raw_q)

#     # detect section (try exact then fuzzy)
#     section = find_section_in_query(raw_q)
#     # If user didn't mention a section, and only one sheet exists, default to it
#     if not section and len(timetables) == 1:
#         section = next(iter(timetables.keys()))

#     day = find_day_in_query(raw_q)
#     hour_label = find_hour_label_in_query(raw_q)
#     wants_incharge = detect_incharge_request(raw_q)
#     wants_free = detect_free_request(raw_q)
#     wants_next = detect_next_request(raw_q)

#     responses = []

#     # If user asked "who is class incharge" (generalized)
#     if wants_incharge:
#         if not section:
#             # try to guess section from query tokens like "2a"
#             sec_guess = find_section_in_query(qnorm)
#             if sec_guess:
#                 section = sec_guess
#         if section:
#             ci = timetables.get(section, {}).get("class_incharge")
#             responses.append(f"👨‍🏫 Class Incharge for {section.upper()}: {ci if ci else 'Not found'}")
#         else:
#             # return incharges for all sections
#             for sec, data in timetables.items():
#                 responses.append(f"👨‍🏫 {sec.upper()}: {data.get('class_incharge')}")
#         return responses

#     # If user requested next/upcoming class
#     if wants_next:
#         if not section:
#             # try to infer; if only one section, use it
#             if len(timetables) == 1:
#                 section = next(iter(timetables.keys()))
#             else:
#                 # try to extract something like "for 2a" in query
#                 section = find_section_in_query(qnorm)
#         if not section:
#             return ["⚠️ Please specify which section (e.g., 'CSE-2A next class')."]
#         # use timezone Asia/Kolkata
#         if ZoneInfo:
#             now = datetime.now(ZoneInfo("Asia/Kolkata"))
#         else:
#             now = datetime.now()
#         today_name = now.strftime("%A")  # 'Monday'
#         schedule = timetables[section]["schedule"]
#         # find today's schedule (fall back to any matching day name ignoring case)
#         day_schedule = None
#         for d in schedule.keys():
#             if d.lower().startswith(today_name.lower()[:3]) or d.lower() == today_name.lower():
#                 day_schedule = schedule[d]
#                 break
#         if not day_schedule:
#             return [f"⚠️ No timetable found for {section.upper()} on {today_name}."]
#         # find the earliest hour whose start time is after now
#         now_t = now.time()
#         upcoming = None
#         for h_label in HOUR_LABELS:
#             start = HOUR_STARTS.get(h_label)
#             if not start:
#                 continue
#             if start > now_t:
#                 # check if slot exists and not 'free'
#                 subj = day_schedule.get(h_label)
#                 upcoming = (h_label, subj, start)
#                 break
#         if upcoming:
#             h_label, subj, start = upcoming
#             return [f"📣 Next class for {section.upper()} today ({today_name}): {h_label} at {start.strftime('%I:%M %p')} -> {subj}"]
#         else:
#             return [f"✅ No more scheduled classes for {section.upper()} today ({today_name})."]
    
#     # If user asked for free slots / break times
#     if wants_free:
#         if not section:
#             section = find_section_in_query(qnorm)
#             if not section and len(timetables) == 1:
#                 section = next(iter(timetables.keys()))
#             if not section:
#                 return ["⚠️ Please specify section to check free slots (e.g., 'Free slots for CSE-2A Friday')."]
#         schedule = timetables[section]["schedule"]
#         target_day = day
#         if not target_day:
#             # if no day mentioned, return free slots across week
#             lines = [f"🕒 Free slots for {section.upper()} (weekly):"]
#             for d, hours in schedule.items():
#                 free = [h for h,s in hours.items() if isinstance(s, str) and "free" in s.lower()]
#                 lines.append(f"{d}: {', '.join(free) if free else 'None'}")
#             return lines
#         # day provided
#         # find matching day key in schedule
#         day_key = None
#         for d in schedule.keys():
#             if d.lower().startswith(target_day.lower()[:3]) or target_day.lower() in d.lower():
#                 day_key = d
#                 break
#         if not day_key:
#             return [f"⚠️ Could not find day '{target_day}' in timetable for {section.upper()}."]
#         free = [h for h,s in schedule[day_key].items() if isinstance(s, str) and "free" in s.lower()]
#         return [f"🕒 Free slots in {section.upper()} on {day_key}: {', '.join(free) if free else 'None'}"]

#     # If user explicitly asked for full timetable or day-specific or hour-specific
#     if "timetable" in qnorm or "schedule" in qnorm or day or hour_label:
#         if not section:
#             # if only one section, default
#             if len(timetables) == 1:
#                 section = next(iter(timetables.keys()))
#             else:
#                 section = find_section_in_query(qnorm)
#         if not section:
#             return ["⚠️ Please specify section (e.g., 'CSE-2A timetable')."]
#         schedule = timetables[section]["schedule"]

#         # full timetable requested (no day)
#         if ("timetable" in qnorm or "schedule" in qnorm) and not day:
#             lines = [f"📅 Timetable for {section.upper()}:"]
#             for d, hours in schedule.items():
#                 lines.append(f"\n--- {d} ---")
#                 for h, sub in hours.items():
#                     lines.append(f"{h}: {sub}")
#             return lines

#         # day-specific
#         if day and not hour_label:
#             # find actual key for day
#             day_key = None
#             for d in schedule.keys():
#                 if d.lower().startswith(day.lower()[:3]) or day.lower() in d.lower():
#                     day_key = d
#                     break
#             if not day_key:
#                 return [f"⚠️ Could not find day '{day}' in timetable for {section.upper()}."]
#             lines = [f"📅 {section.upper()} - {day_key} Timetable:"]
#             for h, sub in schedule[day_key].items():
#                 lines.append(f"{h}: {sub}")
#             return lines

#         # hour-specific
#         if day and hour_label:
#             # find day key
#             day_key = None
#             for d in schedule.keys():
#                 if d.lower().startswith(day.lower()[:3]) or day.lower() in d.lower():
#                     day_key = d
#                     break
#             if not day_key:
#                 return [f"⚠️ Could not find day '{day}' in timetable for {section.upper()}."]
#             subject = schedule[day_key].get(hour_label)
#             return [f"{section.upper()} - {day_key} {hour_label}: {subject}"]

#     # Teacher-based query (when user mentions a person's name or asks "when does <name> teach")
#     teacher_candidates = find_teacher_slots(qnorm, section=section, day=day)
#     if teacher_candidates:
#         # present unique teacher phrases and their slots
#         # We will try to cluster by raw text matches in cells
#         out = []
#         # take top matches and aggregate by teacher raw substrings
#         aggregated = {}
#         for sec, d, h, raw, score in teacher_candidates:
#             key = normalize_text(raw)
#             aggregated.setdefault(key, []).append((sec, d, h, raw, score))
#         for k, slots in aggregated.items():
#             # pick a display name (best raw)
#             display = slots[0][3]
#             out.append(f"👩‍🏫 Matches for '{display}':")
#             for sec,d,h,raw,sc in slots:
#                 out.append(f"- {sec.upper()} {d} {h}: {raw}")
#         return out

#     # Subject-based search
#     subj_slots = find_subject_slots(qnorm, section=section, day=day)
#     if subj_slots:
#         out = []
#         for sec,d,h,raw,score in subj_slots[:10]:
#             out.append(f"{sec.upper()} - {d} {h}: {raw}")
#         return out

#     # fallback: try fuzzy section+day detection and show day timetable
#     if not section:
#         section = find_section_in_query(qnorm)
#     if section and day:
#         schedule = timetables[section]["schedule"]
#         day_key = None
#         for d in schedule.keys():
#             if d.lower().startswith(day.lower()[:3]) or day.lower() in d.lower():
#                 day_key = d
#                 break
#         if day_key:
#             lines = [f"📅 {section.upper()} - {day_key} Timetable:"]
#             for h, sub in schedule[day_key].items():
#                 lines.append(f"{h}: {sub}")
#             return lines

#     return ["🤖 Sorry, I couldn't understand your question. Try: 'CSE-2A Monday timetable', 'Who is class incharge for 2A?', 'When does Mamatha Jain teach?', 'Next class for CSE-2A'."]

# # ----------------- Example quick test (only runs when file executed) -----------------
# if __name__ == "__main__":
#     tests = [
#         "CSE-2A timetable",
#         "2a Monday classes",
#         "Who is incharge for 2A?",
#         "Class teacher cse2a?",
#         "CSE 2A H4 Thursday",
#         "When does mamatha jain teach?",
#         "Next class for cse2a",
#         "Free slots Friday cse2a",
#         "When is chemistry for 2a?",
#         "Who handles 2A?"        
#     ]
#     for t in tests:
#         print("\nQuery:", t)
#         res = gui_get_timetable(t)
#         print("\n".join(res))
#         print("-"*50)
"""
timetable.py — Hybrid NLP-style Timetable Engine (A+B+C)
Features:
- Exact + fuzzy matching for sections, days, hours, teachers, subjects
- Class incharge extraction from A10
- Teacher-based search, subject search, free-slot finder
- Next-class finder (uses Asia/Kolkata timezone)
- Works with your existing Excel layout (A4:A8 days, B-J hours, A10 class incharge)
- NOW ADDED: HELP MENU (no previous features removed)
"""

import pandas as pd
import re
import difflib
from datetime import datetime, time
try:
    from zoneinfo import ZoneInfo  # Python 3.9+
except Exception:
    ZoneInfo = None


# ---------------- CONFIG: map hour labels to start times ----------------
HOUR_STARTS = {
    "H1": time(9, 0),
    "H2": time(9, 50),
    "H3": time(10, 40),
    "H4": time(11, 30),
    "H0": time(12, 20),  # Lunch
    "H5": time(13, 10),
    "H6": time(14, 0),
    "H7": time(14, 50),
    "H8": time(15, 40),
}


# ---------------- HELP TEXT (NEW) ----------------
HELP_TEXT = """
📘 **Timetable Help – Supported Queries**

==============================
🔹 Section Timetable
CSE-2A timetable
Show timetable for 2a
Schedule CSE2A

==============================
🔹 Day Timetable
CSE-2A Monday
2a Friday classes
CSE2A Tuesday timetable

==============================
🔹 Hour-wise Query
cse2a H4 Thursday


==============================
🔹 Class Incharge
Who is class incharge for cse2a?
2a class teacher
class coordinator cse-2a

==============================
🔹 Next Class
next class for cse2a
what is next for cse-2a?
upcoming class-2A

==============================
🔹 Free Slot Finder
free slots cse-2a
free periods for cse-2a on monday
break time friday cse-2a

==============================
🔹 Teacher Search
when does mamatha jain teach?
classes handled by ramesh
mamatha monday

==============================
🔹 Subject Search
when is chemistry for cse-2a?
math  cse-2a
2a physics class

==============================
🔹 Fuzzy Search Supported
csee-2aa
cse-2a
aiml2b
mondey
chemestree

==============================
🤖 Ask naturally!
Example: "What is cse2a 2nd hour on Friday?"
"""


# ---------------- HELP DETECTOR (NEW) ----------------
def detect_help_request(query: str):
    q = normalize_text(query)
    help_words = ["help", "sample", "usage", "queries", "how to", "what can you", "commands"]
    return any(w.replace(" ", "") in q.replace(" ", "") for w in help_words)


# ---------------- Utility helpers ----------------
def normalize_text(s: str) -> str:
    if s is None:
        return ""
    s = str(s).lower()
    s = re.sub(r'[\u200b-\u200d\ufeff]', '', s)
    s = re.sub(r'[\s\-_\./,]+', ' ', s)
    s = re.sub(r'[^a-z0-9 ]', '', s)
    return s.strip()

def simple_tokens(s: str):
    return [t for t in normalize_text(s).split() if t]

def fuzzy_score(a: str, b: str) -> float:
    return difflib.SequenceMatcher(None, a, b).ratio() if a and b else 0.0

def fuzzy_in(text: str, candidate: str, threshold=0.75) -> bool:
    text_n = normalize_text(text)
    candidate_n = normalize_text(candidate)
    if candidate_n in text_n:
        return True
    if fuzzy_score(text_n, candidate_n) >= threshold:
        return True
    words = text_n.split()
    if len(candidate_n.split()) <= 1:
        return any(fuzzy_score(w, candidate_n) >= threshold for w in words)
    m = len(candidate_n.split())
    for i in range(len(words) - m + 1):
        if fuzzy_score(" ".join(words[i:i+m]), candidate_n) >= threshold:
            return True
    return False


# ----------------- Load Excel -----------------
FILE_PATH = "data/timetable.xlsx" #r"C:\Users\Administrator\Desktop\AU\Chatbot\timetable.xlsx"
xls = pd.ExcelFile(FILE_PATH)

timetables = {}
_subject_index = {}
_teacher_index = {}
_all_cells = []

DAY_NAMES = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday"]
DAY_ABBRS = {d[:3]: d for d in DAY_NAMES}

HOUR_LABELS = ['H1','H2','H3','H4','H0','H5','H6','H7','H8']
HOUR_COLS = [1,2,3,4,5,6,7,8,9]


# ----------------- Parse all sheets -----------------
for sheet in xls.sheet_names:
    df = pd.read_excel(FILE_PATH, sheet_name=sheet, header=None)
    section = normalize_text(sheet).replace(" ", "")

    try:
        days = [str(x).strip() for x in df.iloc[3:8, 0].tolist()]
    except:
        days = []

    timetable = {}

    for i, day in enumerate(days):
        day_name = day if isinstance(day, str) and day.strip() else DAY_NAMES[i]
        timetable[day_name] = {}

        for col_idx, hour_label in zip(HOUR_COLS, HOUR_LABELS):
            try:
                raw_cell = df.iloc[i+3, col_idx]
            except:
                raw_cell = None

            timetable[day_name][hour_label] = raw_cell
            _all_cells.append((section, day_name, hour_label, raw_cell))

            if isinstance(raw_cell, str) and raw_cell.strip():
                cell_norm = normalize_text(raw_cell)

                primary = re.split(r',', raw_cell)[0]
                subkey = normalize_text(primary)
                if subkey:
                    _subject_index.setdefault(subkey, []).append((section, day_name, hour_label, raw_cell))

                tokens = cell_norm.split()
                for n in range(1, min(4, len(tokens))+1):
                    for st in range(len(tokens)-n+1):
                        gram = " ".join(tokens[st:st+n])
                        if len(gram) >= 2:
                            _subject_index.setdefault(gram, []).append((section, day_name, hour_label, raw_cell))

                raw_tokens = [t.strip() for t in re.split(r'[,/();\-]', raw_cell) if t.strip()]
                for chunk in raw_tokens:
                    ck = normalize_text(chunk)
                    if len(ck) > 2:
                        _teacher_index.setdefault(ck, []).append(
                            (section, day_name, hour_label, raw_cell)
                        )

    class_incharge = None
    try:
        cell_val = df.iloc[9, 0]
        if isinstance(cell_val, str) and "class incharge" in cell_val.lower():
            class_incharge = cell_val.split(":",1)[-1].strip()
        elif isinstance(cell_val, str) and cell_val.strip():
            class_incharge = cell_val.strip()
    except:
        class_incharge = None

    timetables[section] = {
        "class_incharge": class_incharge,
        "schedule": timetable
    }


# ---------------- Section detector ----------------
def find_section_in_query(query: str):
    q = normalize_text(query).replace(" ", "")
    if q in timetables:
        return q
    keys = list(timetables.keys())
    best = difflib.get_close_matches(q, keys, n=1, cutoff=0.6)
    if best:
        return best[0]
    for k in keys:
        if k in q or q in k:
            return k
    return None


# ---------------- Day detector ----------------
def find_day_in_query(query: str):
    q = normalize_text(query)
    for d in DAY_NAMES:
        if d in q:
            return d
    for ab,fu in DAY_ABBRS.items():
        if ab in q:
            return fu
    for d in DAY_NAMES:
        if fuzzy_score(q, d) >= 0.6:
            return d
    return None


# ---------------- Hour detector ----------------
def find_hour_label_in_query(query: str):
    q = normalize_text(query)
    m = re.search(r'\b(h|hour|hr)?\s*([1-8])\b', query.lower())
    if m:
        num = int(m.group(2))
        label = f"H{num}" if num != 0 else "H0"
        if label in HOUR_LABELS:
            return label
    ord_map = {
        "first":"H1","second":"H2","third":"H3","fourth":"H4",
        "fifth":"H5","sixth":"H6","seventh":"H7","eighth":"H8"
    }
    for word,lab in ord_map.items():
        if word in q:
            return lab
    if "lunch" in q or "break" in q:
        return "H0"
    for lab in HOUR_LABELS:
        if lab.lower() in q:
            return lab
    return None


# ---------------- Flags ----------------
def detect_incharge_request(query: str):
    q = normalize_text(query)
    keys=["incharge","in charge","class teacher","coordinator","mentor"]
    return any(k.replace(" ","") in q.replace(" ","") for k in keys)

def detect_free_request(query: str):
    q = normalize_text(query)
    return any(w in q for w in ["free","free slot","free slots","break","vacant"])

def detect_next_request(query: str):
    q = normalize_text(query)
    return any(w in q for w in ["next","upcoming","now","current","what is next","whats next","next class"])


# ---------------- Teacher search ----------------
def find_teacher_slots(query, section=None, day=None, top_n=20):
    q = normalize_text(query)
    cands = []

    for key,slots in _teacher_index.items():
        score = fuzzy_score(q, key)
        if score >= 0.55 or key in q:
            for s in slots:
                sec,d,h,raw = s
                if section and sec != section:
                    continue
                if day and d.lower() != day.lower():
                    continue
                cands.append((sec,d,h,raw,score))

    if not cands:
        for sec,d,h,raw in _all_cells:
            if isinstance(raw,str):
                if fuzzy_in(raw, q, 0.6):
                    cands.append((sec,d,h,raw,fuzzy_score(normalize_text(raw),q)))

    return sorted(cands, key=lambda x: x[4], reverse=True)[:top_n]


# ---------------- Subject search ----------------
def find_subject_slots(query, section=None, day=None, top_n=10):
    q = normalize_text(query)
    results=[]
    for key,slots in _subject_index.items():
        if fuzzy_score(q,key)>=0.6 or key in q or q in key:
            for s in slots:
                sec,d,h,raw = s
                if section and sec!=section:
                    continue
                if day and d.lower()!=day.lower():
                    continue
                results.append((sec,d,h,raw,fuzzy_score(q,key)))
    return sorted(results, key=lambda x:x[4], reverse=True)[:top_n]


# ---------------- MAIN ENGINE ----------------
def gui_get_timetable(query: str):
    if not query or not query.strip():
        return ["⚠️ Please type a query, e.g., 'CSE-2A Monday timetable'"]

    # === HELP (NEW) ===
    if detect_help_request(query):
        return HELP_TEXT.split("\n")

    raw_q = query.strip()
    qnorm = normalize_text(raw_q)

    section = find_section_in_query(raw_q)
    if not section and len(timetables)==1:
        section = next(iter(timetables.keys()))

    day = find_day_in_query(raw_q)
    hour = find_hour_label_in_query(raw_q)

    wants_incharge = detect_incharge_request(raw_q)
    wants_free = detect_free_request(raw_q)
    wants_next = detect_next_request(raw_q)

    # CLASS INCHARGE
    if wants_incharge:
        if section:
            ci = timetables.get(section,{}).get("class_incharge")
            return [f"👨‍🏫 Class Incharge for {section.upper()}: {ci}"]
        else:
            lines=[]
            for sec,data in timetables.items():
                lines.append(f"{sec.upper()}: {data.get('class_incharge')}")
            return lines

    # NEXT CLASS
    if wants_next:
        if not section:
            if len(timetables)==1:
                section = next(iter(timetables.keys()))
            else:
                return ["⚠️ Please specify which section for next class."]

        if ZoneInfo:
            now = datetime.now(ZoneInfo("Asia/Kolkata"))
        else:
            now = datetime.now()

        today = now.strftime("%A")
        schedule = timetables[section]["schedule"]

        day_sched=None
        for d in schedule.keys():
            if d.lower().startswith(today.lower()[:3]):
                day_sched = schedule[d]
        if not day_sched:
            return [f"No timetable found for {today}"]

        now_t=now.time()
        for h in HOUR_LABELS:
            st = HOUR_STARTS.get(h)
            if st and st>now_t:
                return [f"📣 Next class for {section.upper()} ({today}): {h} → {day_sched[h]}"]

        return [f"✅ No more classes today for {section.upper()}"]

    # FREE SLOT
    if wants_free:
        if not section:
            return ["⚠️ Mention section, e.g., 'free slots cse2a'"]

        schedule=timetables[section]["schedule"]

        if not day:
            lines=[f"🕒 Free slots (weekly) – {section.upper()}"]
            for d,hrs in schedule.items():
                free=[h for h,s in hrs.items() if isinstance(s,str) and "free" in s.lower()]
                lines.append(f"{d}: {', '.join(free) if free else 'None'}")
            return lines

        dk=None
        for d in schedule.keys():
            if day in d.lower():
                dk=d
        if not dk:
            return [f"Day '{day}' not found"]

        free=[h for h,s in schedule[dk].items() if isinstance(s,str) and "free" in s.lower()]
        return [f"🕒 Free slots on {dk} for {section.upper()}: {', '.join(free)}"]

    # TIMETABLE / DAY / HOUR HANDLING
    if "timetable" in qnorm or "schedule" in qnorm or day or hour:
        if not section:
            return ["⚠️ Which section? e.g., 'CSE-2A timetable'"]

        schedule=timetables[section]["schedule"]

        # Full section timetable
        if ("timetable" in qnorm or "schedule" in qnorm) and not day:
            lines=[f"📅 Timetable for {section.upper()}"]
            for d,hrs in schedule.items():
                lines.append(f"\n--- {d} ---")
                for h,s in hrs.items():
                    lines.append(f"{h}: {s}")
            return lines

        # Day timetable
        if day and not hour:
            dk=None
            for d in schedule.keys():
                if day[:3] == d[:3].lower():
                    dk=d
            if not dk:
                return [f"Day '{day}' not found"]
            lines=[f"📅 {section.upper()} - {dk}"]
            for h,s in schedule[dk].items():
                lines.append(f"{h}: {s}")
            return lines

        # hour-specific
        if day and hour:
            dk=None
            for d in schedule.keys():
                if day[:3]==d[:3].lower():
                    dk=d
            if not dk:
                return [f"Day '{day}' not found"]

            subject=schedule[dk].get(hour)
            return [f"{section.upper()} - {dk} {hour}: {subject}"]

    # TEACHER SEARCH
    teacher = find_teacher_slots(qnorm, section=section, day=day)
    if teacher:
        out=[]
        for sec,d,h,raw,score in teacher:
            out.append(f"{sec.upper()} - {d} {h}: {raw}")
        return out

    # SUBJECT SEARCH
    subj = find_subject_slots(qnorm, section=section, day=day)
    if subj:
        out=[]
        for sec,d,h,raw,score in subj:
            out.append(f"{sec.upper()} - {d} {h}: {raw}")
        return out

    return ["🤖 Sorry, I couldn't understand. Try 'timetable help'."]
