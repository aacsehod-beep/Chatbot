# import pandas as pd
# from difflib import get_close_matches

# # ----------------- Load Excel -----------------
# file_path = r"C:\Users\Administrator\Desktop\AU\Chatbot\student_contacts.xlsx"
# all_sheets = pd.read_excel(file_path, sheet_name=None)

# # Add section name as a column
# for sheet_name, df in all_sheets.items():
#     df["Section"] = sheet_name

# # Merge all sheets
# data = pd.concat(all_sheets.values(), ignore_index=True)
# data.columns = data.columns.str.strip()

# # ----------------- Function -----------------
# def get_student_info(query: str):
#     query_lower = query.strip().lower()
#     if not query_lower:
#         return "❌ Please enter a valid query."

#     # Detect section name if provided
#     sections = data['Section'].str.lower().unique().tolist()
#     section = None
#     for sec in sections:
#         if sec in query_lower:
#             section = sec
#             query_lower = query_lower.replace(sec, "").strip()
#             break

#     # Detect request type
#     want_parent = "parent" in query_lower and "no" in query_lower
#     want_contact = "contact" in query_lower and "no" in query_lower

#     # Clean query
#     clean_query = (
#         query_lower.replace("parent", "")
#         .replace("contact", "")
#         .replace("no", "")
#         .strip()
#     )

#     # Choose dataset
#     search_df = data if not section else data[data['Section'].str.lower() == section]

#     # ----------------- Matching -----------------
#     matched = search_df[search_df['Name'].str.lower().str.contains(clean_query, na=False)]

#     if matched.empty and clean_query:
#         all_names = search_df['Name'].dropna().str.lower().tolist()
#         close = get_close_matches(clean_query, all_names, n=5, cutoff=0.7)
#         if close:
#             matched = search_df[search_df['Name'].str.lower().isin(close)]

#     if matched.empty:
#         if section:
#             return f"❌ No student found in section {section.upper()} for: {clean_query}"
#         else:
#             return f"❌ No student found for: {clean_query} in any section."

#     # ----------------- Multiple Matches -----------------
#     responses = []
#     for _, student in matched.iterrows():
#         student = student.to_dict()

#         if want_parent:
#             responses.append(
#                 f"📞 Parent Contact for {student['Name']} ({student['Section']}): {student['Parent Contact No']}"
#             )
#         elif want_contact:
#             responses.append(
#                 f"📱 Student Contact for {student['Name']} ({student['Section']}): {student['Student Contact No']}"
#             )
#         else:
#             info = (
#                 f"\n📌 Student Information\n"
#                 f"-----------------------\n"
#                 f"👤 Name: {student['Name']}\n"
#                 f"🆔 Reg. No: {student['Reg.No']}\n"
#                 f"📱 Student Contact: {student['Student Contact No']}\n"
#                 f"📞 Parent Contact: {student['Parent Contact No']}\n"
#                 f"✉️ Email: {student['Student official mail id']}\n"
#                 f"🏫 Section: {student['Section']}\n"
#             )
#             responses.append(info)

#     # ----------------- Final Output -----------------
#     if len(responses) > 1:
#         header = "⚠️ Multiple students found:\n"
#     else:
#         header = ""
#     return header + "\n".join(responses)
"""
# Enhanced student info module
# Features:
# - Loads all sheets from student_contacts.xlsx and normalizes columns
# - Builds fast indexes (name, regno, student contact, parent contact, section)
# - Fuzzy/partial/name-order matching (get_close_matches)
# - Intent detection for many query types:
#     * parent contact, student contact
#     * email, reg no, full info
#     * list students in a section
#     * search by reg no (exact or partial)
#     * search by phone (student/parent)
#     * generic name lookup
# - Clean, chatbot-friendly output
# - Example usage in __main__
# """

# import pandas as pd
# import re
# from difflib import get_close_matches
# from typing import List, Dict, Any, Optional

# # ---------- CONFIG ----------
# FILE_PATH = r"C:\Users\Administrator\Desktop\AU\Chatbot\student_contacts.xlsx"
# FUZZY_CUTOFF = 0.7  # 0-1 (higher = stricter)
# MAX_FUZZY_MATCHES = 6

# # ---------- LOADING & NORMALIZATION ----------
# _raw_sheets: Dict[str, pd.DataFrame] = pd.read_excel(FILE_PATH, sheet_name=None)

# # Standardize column names we expect (lowered and trimmed)
# def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
#     df = df.copy()
#     df.columns = [str(c).strip() for c in df.columns]
#     # common alternative column names mapping
#     col_map = {}
#     for col in df.columns:
#         lc = col.lower()
#         if "name" in lc and "parent" not in lc:
#             col_map[col] = "Name"
#         elif ("reg" in lc and "no" in lc) or ("register" in lc):
#             col_map[col] = "Reg.No"
#         elif "student" in lc and ("contact" in lc or "phone" in lc or "mobile" in lc):
#             col_map[col] = "Student Contact No"
#         elif "parent" in lc and ("contact" in lc or "phone" in lc or "mobile" in lc):
#             col_map[col] = "Parent Contact No"
#         elif "mail" in lc or "email" in lc:
#             col_map[col] = "Student official mail id"
#         else:
#             # keep it, but capitalized nicely
#             col_map[col] = col
#     df = df.rename(columns=col_map)
#     return df

# _sheets = {}
# for sheet_name, df in _raw_sheets.items():
#     df = _normalize_columns(df)
#     # add section column if not present
#     if "Section" not in df.columns:
#         df["Section"] = sheet_name
#     else:
#         # if Section exists but empty, fill with sheet_name where missing
#         df["Section"] = df["Section"].fillna(sheet_name)
#     _sheets[sheet_name] = df

# # Concatenate into a single DataFrame for searching
# data = pd.concat(_sheets.values(), ignore_index=True, sort=False)
# # Ensure columns exist
# for col in ["Name", "Reg.No", "Student Contact No", "Parent Contact No", "Student official mail id", "Section"]:
#     if col not in data.columns:
#         data[col] = pd.NA

# # Trim whitespace and create lowercased helper columns
# data["Name_clean"] = data["Name"].astype(str).str.strip()
# data["Name_lower"] = data["Name_clean"].str.lower()
# data["Reg_clean"] = data["Reg.No"].astype(str).str.strip()
# data["Section_clean"] = data["Section"].astype(str).str.strip()
# data["Section_lower"] = data["Section_clean"].str.lower()
# data["Student_contact_clean"] = data["Student Contact No"].astype(str).str.strip()
# data["Parent_contact_clean"] = data["Parent Contact No"].astype(str).str.strip()
# data["Email_clean"] = data["Student official mail id"].astype(str).str.strip()

# # ---------- INDEXES ----------
# # name -> list of indices
# _name_to_indices: Dict[str, List[int]] = {}
# for idx, name in enumerate(data["Name_lower"]):
#     if pd.isna(name) or str(name).strip() == "nan":
#         continue
#     _name_to_indices.setdefault(name, []).append(idx)

# # regno -> index
# _reg_to_index: Dict[str, int] = {}
# for idx, reg in enumerate(data["Reg_clean"]):
#     if pd.isna(reg) or str(reg).strip() == "nan":
#         continue
#     _reg_to_index[reg] = idx

# # phone -> indices (both student and parent)
# _phone_to_indices: Dict[str, List[int]] = {}
# for idx, phone in enumerate(data["Student_contact_clean"]):
#     if pd.isna(phone) or str(phone).strip() == "nan":
#         continue
#     p = re.sub(r"\D", "", str(phone))
#     if p:
#         _phone_to_indices.setdefault(p, []).append(idx)

# for idx, phone in enumerate(data["Parent_contact_clean"]):
#     if pd.isna(phone) or str(phone).strip() == "nan":
#         continue
#     p = re.sub(r"\D", "", str(phone))
#     if p:
#         _phone_to_indices.setdefault(p, []).append(idx)

# # section -> list of indices
# _section_to_indices: Dict[str, List[int]] = {}
# for idx, sec in enumerate(data["Section_lower"]):
#     if pd.isna(sec) or str(sec).strip() == "nan":
#         continue
#     _section_to_indices.setdefault(sec, []).append(idx)

# # prebuilt name list for fuzzy matches
# _all_names_lower = sorted(set([n for n in data["Name_lower"].dropna().tolist() if n and n != "nan"]))

# # ---------- UTILITIES ----------
# def _format_student_row(idx: int) -> str:
#     row = data.iloc[idx]
#     parts = [
#         f"👤 Name: {row['Name_clean']}",
#         f"🏷️ Reg.No: {row['Reg_clean']}",
#         f"🏫 Section: {row['Section_clean']}",
#         f"📱 Student Contact: {row['Student_contact_clean']}",
#         f"📞 Parent Contact: {row['Parent_contact_clean']}",
#         f"✉️ Email: {row['Email_clean']}",
#     ]
#     return "\n".join(parts)

# def _unique_indices_from_list(idxs: List[int]) -> List[int]:
#     seen = set()
#     out = []
#     for i in idxs:
#         if i not in seen:
#             seen.add(i)
#             out.append(i)
#     return out

# def _fuzzy_name_match(name_query: str) -> List[int]:
#     """
#     Try several strategies:
#     - direct contains match
#     - initials matching
#     - reversed name parts
#     - difflib.get_close_matches on full names
#     Returns list of indices (may be empty).
#     """
#     q = name_query.strip().lower()
#     if not q:
#         return []

#     matched_indices: List[int] = []

#     # 1) direct contains in Name_lower
#     contains_mask = data["Name_lower"].str.contains(re.escape(q), na=False)
#     matched_indices.extend(list(data[contains_mask].index))

#     # 2) initials matching (e.g., "b s" -> matches "Bharath S")
#     tokens = q.split()
#     if len(tokens) >= 2 and all(len(t) == 1 for t in tokens):
#         # build regex for initials e.g. ^b.*\s+s.*$
#         pattern = r"^" + r"\s*".join([t + r"[^\s]*" for t in tokens])
#         initials_mask = data["Name_lower"].str.match(pattern, na=False)
#         matched_indices.extend(list(data[initials_mask].index))

#     # 3) reversed name (e.g., "kumar bharath")
#     rev_tokens = " ".join(reversed(tokens))
#     if rev_tokens != q:
#         rev_mask = data["Name_lower"].str.contains(re.escape(rev_tokens), na=False)
#         matched_indices.extend(list(data[rev_mask].index))

#     # 4) fuzzy close matches on full names list
#     if not matched_indices:
#         # try get_close_matches on the list of full names
#         close = get_close_matches(q, _all_names_lower, n=MAX_FUZZY_MATCHES, cutoff=FUZZY_CUTOFF)
#         for c in close:
#             matched_indices.extend(_name_to_indices.get(c, []))

#     return _unique_indices_from_list(matched_indices)

# # ---------- INTENT DETECTION ----------
# def _detect_intent_and_extractor(query: str) -> Dict[str, Any]:
#     """
#     Returns dict with possible keys:
#       - intent: one of:
#         fetch_parent, fetch_student_contact, fetch_email, fetch_reg_no,
#         fetch_full_info, list_section, search_by_reg, search_by_phone, search_by_name
#       - value: the extracted name/reg/phone/section/...
#     """
#     q = query.strip()
#     q_lower = q.lower()

#     # Quick helpers
#     digits = re.sub(r"\D", "", q)
#     tokens = q_lower.split()

#     # 1) Section listing: "list students in AIML-2B", "show all in cse a"
#     m = re.search(r"\b(list|show|give)\b.*\b(students|students list|students in|all)\b.*\b(in|of|for)\b\s*(?P<section>[a-z0-9\-\s]+)$", q_lower)
#     if m:
#         sec = m.group("section").strip()
#         return {"intent": "list_section", "value": sec}

#     # simpler "students in AIML-2B" or "aiml-2b students"
#     m = re.search(r"\b(?P<section>[a-z0-9\-]+\s*[a-z0-9\-]*)\b.*\bstudents\b", q_lower)
#     if m and len(m.group("section").strip()) <= 30:
#         sec = m.group("section").strip()
#         # confirm it looks like a section (exists in our sections)
#         if sec in _section_to_indices:
#             return {"intent": "list_section", "value": sec}

#     # 2) Phone number lookup (user gave phone): exact digits search
#     if len(digits) >= 6:
#         # Could be reg no or phone. We'll heuristically check:
#         # If digits length >= 9 treat as phone first
#         if len(digits) >= 9:
#             return {"intent": "search_by_phone", "value": digits}
#         # If digits looks like a regno pattern (contains letters earlier) prefer reg search below
#         # We'll still fall through to search-by-reg if pattern matches

#     # 3) Reg no exact or partial: typical regno contains letters + digits
#     # match patterns like 21A11A05G2 or last-4 digits etc.
#     reg_candidate = re.search(r"([A-Za-z0-9]{6,})", q.replace(" ", ""))
#     if reg_candidate:
#         cand = reg_candidate.group(1)
#         # if the candidate matches an existing reg exactly
#         if cand in _reg_to_index:
#             return {"intent": "search_by_reg", "value": cand}
#         # partial reg: user might have typed last few characters
#         if len(re.sub(r"\D", "", cand)) >= 3:
#             return {"intent": "search_by_reg", "value": cand}

#     # 4) Intent words for parent/student contact
#     if "parent" in q_lower and ("contact" in q_lower or "no" in q_lower or "number" in q_lower or "phone" in q_lower):
#         # extract name if present (remove 'parent', 'contact')
#         name = re.sub(r"\bparent\b|\bcontact\b|\bnumber\b|\bno\b|\bphone\b", "", q_lower).strip()
#         return {"intent": "fetch_parent", "value": name}

#     if ("student" in q_lower and ("contact" in q_lower or "phone" in q_lower or "mobile" in q_lower)) or re.search(r"\b(student phone|student contact|student mobile|mobile of)\b", q_lower):
#         name = re.sub(r"\bstudent\b|\bcontact\b|\bphone\b|\bmobile\b", "", q_lower).strip()
#         return {"intent": "fetch_student_contact", "value": name}

#     # 5) Email intent
#     if "email" in q_lower or "mail" in q_lower:
#         name = re.sub(r"\b(email|mail|id|what is)\b", "", q_lower).strip()
#         return {"intent": "fetch_email", "value": name}

#     # 6) Reg no ask (explicit)
#     if "reg" in q_lower or "registration" in q_lower or "register" in q_lower:
#         name = re.sub(r"\b(reg|reg\.|regno|registration|register|registration no|register no)\b", "", q_lower).strip()
#         return {"intent": "fetch_reg_no", "value": name}

#     # 7) Generic "who is <name>" or "info about <name>" -> full info
#     m = re.search(r"\b(info|information|details|who is|tell me about|show)\b\s*(?P<name>.+)", q_lower)
#     if m:
#         nm = m.group("name").strip()
#         return {"intent": "fetch_full_info", "value": nm}

#     # 8) If query contains word 'students' plus a section we covered earlier - already handled
#     # 9) Fallback: treat as name lookup
#     return {"intent": "search_by_name", "value": q_lower.strip()}

# # ---------- SEARCH HANDLERS ----------
# def _handle_list_section(section: str) -> str:
#     sec = section.strip().lower()
#     # try exact section match or fuzzy among sections
#     if sec in _section_to_indices:
#         idxs = _section_to_indices[sec]
#     else:
#         # fuzzy match against available sections
#         close = get_close_matches(sec, list(_section_to_indices.keys()), n=1, cutoff=0.6)
#         if close:
#             idxs = _section_to_indices[close[0]]
#         else:
#             return f"❌ Section not found: '{section}'."
#     lines = [f"📚 Students in {data.iloc[i]['Section_clean']}: {data.iloc[i]['Name_clean']} (Reg: {data.iloc[i]['Reg_clean']})" for i in idxs]
#     if not lines:
#         return f"❌ No students found in section {section}."
#     return "📋 " + "\n".join(lines)

# def _handle_search_by_phone(phone_digits: str) -> str:
#     d = re.sub(r"\D", "", str(phone_digits))
#     if not d:
#         return "❌ Please provide a valid phone number or partial digits."
#     # try exact or partial match in phone index keys
#     matched_idxs = []
#     # exact numbers
#     if d in _phone_to_indices:
#         matched_idxs.extend(_phone_to_indices[d])
#     # partial match: find any key that contains the digits
#     for key, idxs in _phone_to_indices.items():
#         if d in key and key != d:
#             matched_idxs.extend(idxs)
#     matched_idxs = _unique_indices_from_list(matched_idxs)
#     if not matched_idxs:
#         return f"❌ No student/parent found with phone matching '{phone_digits}'."
#     lines = [f"{_format_student_row(i)}" for i in matched_idxs]
#     return "\n\n".join(lines)

# def _handle_search_by_reg(reg_query: str) -> str:
#     cand = str(reg_query).strip()
#     if not cand:
#         return "❌ Please provide a registration number or partial reg."
#     # exact
#     if cand in _reg_to_index:
#         idx = _reg_to_index[cand]
#         return _format_student_row(idx)
#     # partial/reg substring search
#     mask = data["Reg_clean"].str.contains(re.escape(cand), na=False, case=False)
#     if mask.any():
#         idxs = list(data[mask].index)
#         lines = [f"{_format_student_row(i)}" for i in idxs]
#         return "\n\n".join(lines)
#     # maybe user gave digits only; try matching last-digits
#     digits = re.sub(r"\D", "", cand)
#     if digits and len(digits) >= 3:
#         mask2 = data["Reg_clean"].astype(str).str.contains(digits, na=False)
#         if mask2.any():
#             idxs = list(data[mask2].index)
#             lines = [f"{_format_student_row(i)}" for i in idxs]
#             return "\n\n".join(lines)
#     return f"❌ No registration number matching '{reg_query}' found."

# def _handle_fetch_parent(name_query: str) -> str:
#     idxs = _resolve_name_to_indices(name_query)
#     if not idxs:
#         return f"❌ No student found for: '{name_query}'"
#     lines = []
#     for i in idxs:
#         p = data.iloc[i]["Parent_contact_clean"]
#         lines.append(f"📞 Parent contact for {data.iloc[i]['Name_clean']} ({data.iloc[i]['Section_clean']}): {p}")
#     return "\n".join(lines)

# def _handle_fetch_student_contact(name_query: str) -> str:
#     idxs = _resolve_name_to_indices(name_query)
#     if not idxs:
#         return f"❌ No student found for: '{name_query}'"
#     lines = []
#     for i in idxs:
#         s = data.iloc[i]["Student_contact_clean"]
#         lines.append(f"📱 Student contact for {data.iloc[i]['Name_clean']} ({data.iloc[i]['Section_clean']}): {s}")
#     return "\n".join(lines)

# def _handle_fetch_email(name_query: str) -> str:
#     idxs = _resolve_name_to_indices(name_query)
#     if not idxs:
#         return f"❌ No student found for: '{name_query}'"
#     lines = []
#     for i in idxs:
#         e = data.iloc[i]["Email_clean"]
#         lines.append(f"✉️ Email for {data.iloc[i]['Name_clean']} ({data.iloc[i]['Section_clean']}): {e}")
#     return "\n".join(lines)

# def _handle_fetch_reg_no(name_query: str) -> str:
#     idxs = _resolve_name_to_indices(name_query)
#     if not idxs:
#         return f"❌ No student found for: '{name_query}'"
#     lines = []
#     for i in idxs:
#         lines.append(f"🏷️ Reg.No for {data.iloc[i]['Name_clean']} ({data.iloc[i]['Section_clean']}): {data.iloc[i]['Reg_clean']}")
#     return "\n".join(lines)

# def _handle_fetch_full_info(name_query: str) -> str:
#     idxs = _resolve_name_to_indices(name_query)
#     if not idxs:
#         return f"❌ No student found for: '{name_query}'"
#     lines = []
#     for i in idxs:
#         lines.append(_format_student_row(i))
#     return "\n\n".join(lines)

# def _handle_search_by_name(name_query: str) -> str:
#     idxs = _resolve_name_to_indices(name_query)
#     if not idxs:
#         return f"❌ No student found for: '{name_query}'"
#     if len(idxs) == 1:
#         return _format_student_row(idxs[0])
#     # multiple
#     lines = [f"⚠️ Multiple matches:\n{_format_student_row(i)}" for i in idxs]
#     return "\n\n".join(lines)

# # ---------- HELPER: resolve name to indices ----------
# def _resolve_name_to_indices(name_query: str) -> List[int]:
#     """
#     Return list of matching indices for a given name-like query.
#     Tries:
#       - exact lower match
#       - contains match
#       - initials/reversed/fuzzy via _fuzzy_name_match
#       - if query is short and matches many, returns those
#     """
#     q = (name_query or "").strip().lower()
#     if not q:
#         return []

#     # if user passed a section qualifier like "bharath cse-a", strip section
#     # detect trailing section tokens
#     # e.g., "bharath cse-a" -> name part = bharath
#     parts = q.split()
#     # if last token matches a section name, drop it for name matching
#     if parts and parts[-1] in _section_to_indices:
#         parts = parts[:-1]
#         q = " ".join(parts).strip()

#     # 1) exact
#     if q in _name_to_indices:
#         return _name_to_indices[q][:]

#     # 2) contains
#     mask = data["Name_lower"].str.contains(re.escape(q), na=False)
#     idxs = list(data[mask].index)
#     if idxs:
#         return _unique_indices_from_list(idxs)

#     # 3) initials/reverse/fuzzy
#     fuzzy_idxs = _fuzzy_name_match(q)
#     if fuzzy_idxs:
#         return fuzzy_idxs

#     # 4) If user provided regno-like or phone-like search, try those
#     digits = re.sub(r"\D", "", q)
#     if digits and len(digits) >= 3:
#         # check reg partial
#         maskr = data["Reg_clean"].astype(str).str.contains(digits, na=False)
#         if maskr.any():
#             return list(data[maskr].index)
#         # check phone partial
#         maskp = data["Student_contact_clean"].astype(str).str.contains(digits, na=False) | data["Parent_contact_clean"].astype(str).str.contains(digits, na=False)
#         if maskp.any():
#             return list(data[maskp].index)
#     return []

# # ---------- PUBLIC API ----------
# def get_student_info(query: str) -> str:
#     """
#     Main entry point. Pass a natural language query string.
#     Returns a string suitable for chatbot reply.
#     """
#     if not query or not str(query).strip():
#         return "❌ Please enter a valid query."

#     intent_data = _detect_intent_and_extractor(query)
#     intent = intent_data.get("intent")
#     value = intent_data.get("value", "").strip()

#     try:
#         if intent == "list_section":
#             return _handle_list_section(value)
#         elif intent == "search_by_phone":
#             return _handle_search_by_phone(value)
#         elif intent == "search_by_reg":
#             return _handle_search_by_reg(value)
#         elif intent == "fetch_parent":
#             return _handle_fetch_parent(value)
#         elif intent == "fetch_student_contact":
#             return _handle_fetch_student_contact(value)
#         elif intent == "fetch_email":
#             return _handle_fetch_email(value)
#         elif intent == "fetch_reg_no":
#             return _handle_fetch_reg_no(value)
#         elif intent == "fetch_full_info":
#             return _handle_fetch_full_info(value)
#         elif intent == "search_by_name":
#             return _handle_search_by_name(value)
#         else:
#             # fallback to name search
#             return _handle_search_by_name(value or query)
#     except Exception as e:
#         return f"❌ An error occurred while searching: {e}"

# # ---------- CLI / Testing ----------
# if __name__ == "__main__":
#     tests = [
#         "Bharath",  # name
#         "bharath parent no",  # parent
#         "student phone of Priya",  # student contact
#         "email of Naveen",  # email
#         "reg no of Suresh",  # reg
#         "list students in AIML-2B",  # section list
#         "9876543210",  # phone search
#         "21A11A05G2",  # reg exact
#         "search by 05G2",  # partial reg
#         "info about varalakshmi",  # full info
#     ]
#     for t in tests:
#         print(">>> QUERY:", t)
#         print(get_student_info(t))
#         print("---\n")
"""
Enhanced Student Info Module (FINAL VERSION WITH NAME+SECTION FILTER)
"""

import pandas as pd
import re
from difflib import get_close_matches
from typing import List, Dict


# ==============================
# CONFIG
# ==============================
FILE_PATH ="data/student_contacts.xlsx"  #r"C:\Users\Administrator\Desktop\AU\Chatbot\student_contacts.xlsx"
FUZZY_CUTOFF = 0.7
MAX_FUZZY_MATCHES = 6


# ==============================
# LOAD & NORMALIZE
# ==============================
_raw_sheets: Dict[str, pd.DataFrame] = pd.read_excel(FILE_PATH, sheet_name=None)

def _normalize_columns(df: pd.DataFrame):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    col_map={}
    for col in df.columns:
        lc=col.lower()

        if "name" in lc and "parent" not in lc:
            col_map[col]="Name"
        elif "reg" in lc and "no" in lc:
            col_map[col]="Reg.No"
        elif "student" in lc and ("phone" in lc or "contact" in lc or "mobile" in lc):
            col_map[col]="Student Contact No"
        elif "parent" in lc and ("phone" in lc or "contact" in lc or "mobile" in lc):
            col_map[col]="Parent Contact No"
        elif "mail" in lc or "email" in lc:
            col_map[col]="Student official mail id"
        else:
            col_map[col]=col

    return df.rename(columns=col_map)


_sheets={}
for sheet,df in _raw_sheets.items():
    df=_normalize_columns(df)
    if "Section" not in df.columns:
        df["Section"]=sheet
    df["Section"]=df["Section"].fillna(sheet)
    _sheets[sheet]=df

data=pd.concat(_sheets.values(), ignore_index=True)

required=["Name","Reg.No","Student Contact No","Parent Contact No","Student official mail id","Section"]
for col in required:
    if col not in data.columns:
        data[col]=pd.NA


# ==============================
# CLEANING (all .str.strip())
# ==============================
data["Name_clean"]=data["Name"].astype(str).str.strip()
data["Name_lower"]=data["Name_clean"].str.lower()

data["Reg_clean"]=data["Reg.No"].astype(str).str.strip()

data["Section_clean"]=data["Section"].astype(str).str.strip()
data["Section_lower"]=data["Section_clean"].str.lower()

data["Student_contact_clean"]=data["Student Contact No"].astype(str).str.strip()
data["Parent_contact_clean"]=data["Parent Contact No"].astype(str).str.strip()
data["Email_clean"]=data["Student official mail id"].astype(str).str.strip()


# ==============================
# INDEXES
# ==============================
_name_to_indices={}
for idx,name in enumerate(data["Name_lower"]):
    if name and name!="nan":
        _name_to_indices.setdefault(name,[]).append(idx)

_reg_to_index={}
for idx,reg in enumerate(data["Reg_clean"]):
    if reg and reg!="nan":
        _reg_to_index[reg]=idx

_phone_to_indices={}
for idx,ph in enumerate(data["Student_contact_clean"]):
    p=re.sub(r"\D","",ph)
    if p: _phone_to_indices.setdefault(p,[]).append(idx)

for idx,ph in enumerate(data["Parent_contact_clean"]):
    p=re.sub(r"\D","",ph)
    if p: _phone_to_indices.setdefault(p,[]).append(idx)

_section_to_indices={}
for idx,sec in enumerate(data["Section_lower"]):
    _section_to_indices.setdefault(sec,[]).append(idx)

_all_names_lower=sorted(set(data["Name_lower"].tolist()))


# ==============================
# HELP TEXT
# ==============================
HELP_TEXT = """
📘 **Student Info – Supported Queries**

=============================
🔹 Search by Name
bharath  
varsha  
sr prya  
brahth (fuzzy)  

=============================
🔹 Student Contact
student contact bharath  
kushal phone no  
priya phone number  

=============================
🔹 Parent Contact
bharath parent no  
parent contact of naveen  

=============================
🔹 Email
email of varsha  
priya mail id  

=============================
🔹 Reg.No Search
21A11A05G2  
241U1R1001  
05G2 (partial reg)  

=============================
🔹 Phone Reverse Lookup
9876543210  
search 3210  

=============================
🔹 Section-wise Listing
cse-2a list  
students in aiml-2b  
show ece 1a  

=============================
🔹 Name + Section Filter
kushal cse-2a  
priya aiml 1b  
bharath ece1a  

=============================
🔹 Full Info
info about bharath  
details of priya  

✨ Natural language queries supported.
"""


# ==============================
# UTILITIES
# ==============================
def _unique(lst):
    s=set(); out=[]
    for x in lst:
        if x not in s:
            s.add(x); out.append(x)
    return out


def _format(idx):
    r=data.iloc[idx]
    return (
        f"👤 Name: {r['Name_clean']}\n"
        f"🏷️ Reg.No: {r['Reg_clean']}\n"
        f"🏫 Section: {r['Section_clean']}\n"
        f"📱 Student Contact: {r['Student_contact_clean']}\n"
        f"📞 Parent Contact: {r['Parent_contact_clean']}\n"
        f"✉️ Email: {r['Email_clean']}"
    )


def _fuzzy_name(q):
    q=q.lower().strip()
    out=[]

    out+=data[data["Name_lower"].str.contains(re.escape(q), na=False)].index.tolist()

    tokens=q.split()
    if len(tokens)>=2 and all(len(x)==1 for x in tokens):
        pattern="^" + r"\s*".join([x+r".*" for x in tokens])
        out+=data[data["Name_lower"].str.match(pattern,na=False)].index.tolist()

    rev=" ".join(reversed(tokens))
    if rev!=q:
        out+=data[data["Name_lower"].str.contains(re.escape(rev),na=False)].index.tolist()

    if not out:
        close=get_close_matches(q,_all_names_lower,n=6,cutoff=FUZZY_CUTOFF)
        for c in close:
            out+=_name_to_indices.get(c,[])

    return _unique(out)


# ==============================
# INTENT DETECTOR
# ==============================
def _detect_intent(q_raw: str):
    q=q_raw.lower().strip()
    digits=re.sub(r"\D","",q)
    alnum=re.sub(r"[^A-Za-z0-9]","",q_raw)

    # HELP
    if any(h in q for h in ["help","sample","usage","what can","queries","how to"]):
        return {"intent":"help"}

    # NAME + SECTION FILTER
    name_sec_pattern = r"([a-z]+)\s+([a-z]{2,4}\s*[-]?\s*\d{1,2}[a-z]?)"
    m = re.search(name_sec_pattern, q)
    if m:
        name_part = m.group(1).strip()
        sec_part = m.group(2).replace(" ", "").lower()
        return {"intent": "name_in_section", "value": (name_part, sec_part)}

    # SECTION-FIRST FIX (cse-2a list)
    section_pattern=r"\b([a-z]{2,4}\s*[-]?\s*\d{1,2}[a-z]?)\b"
    sec_match=re.search(section_pattern,q)
    if sec_match:
        clean_sec=sec_match.group(1).replace(" ","").lower()
        if "list" in q or "student" in q or "students" in q or "show" in q:
            return {"intent":"list_section","value":clean_sec}

    # REGNO FIX (letters + digits → regno)
    if any(c.isalpha() for c in alnum) and any(c.isdigit() for c in alnum):
        return {"intent":"search_by_reg","value":alnum}

    # PHONE (digits only)
    if alnum.isdigit() and len(alnum)>=7:
        return {"intent":"search_by_phone","value":alnum}

    # 'kushal phone no'
    if ("phone" in q or "mobile" in q) and ("no" in q or "number" in q):
        nm=re.sub(r"(phone|mobile|no|number|contact)","",q).strip()
        return {"intent":"fetch_student_contact","value":nm}

    # student contact
    if "student" in q and ("contact" in q or "phone" in q or "mobile" in q):
        nm=re.sub(r"(student|contact|phone|mobile)","",q).strip()
        return {"intent":"fetch_student_contact","value":nm}

    # parent contact
    if "parent" in q and ("contact" in q or "no" in q or "number" in q):
        nm=re.sub(r"(parent|contact|number|no)","",q).strip()
        return {"intent":"fetch_parent","value":nm}

    # email
    if "email" in q or "mail" in q:
        nm=re.sub(r"(email|mail|id|address|what is)","",q).strip()
        return {"intent":"fetch_email","value":nm}

    # full info
    if any(x in q for x in ["details","info","who is"]):
        nm=re.sub(r"(details|info|about|who is)","",q).strip()
        return {"intent":"fetch_full_info","value":nm}

    return {"intent":"search_by_name","value":q}


# ==============================
# SEARCH HANDLERS
# ==============================
def _resolve(name): return _fuzzy_name(name)

def _section_list(sec):
    sec=sec.replace(" ","").lower()

    sec_keys=list(_section_to_indices.keys())
    close=get_close_matches(sec,sec_keys,n=1,cutoff=0.5)
    if not close:
        return f"❌ Section '{sec}' not found."

    sec=close[0]
    idxs=_section_to_indices[sec]

    out=[f"📘 Students in {sec.upper()}:"]
    for i in idxs:
        r=data.iloc[i]
        out.append(f"• {r['Name_clean']} ({r['Reg_clean']})")
    return "\n".join(out)


def _name_in_section(name, sec):
    sec_keys=list(_section_to_indices.keys())
    close=get_close_matches(sec,sec_keys,n=1,cutoff=0.5)
    if not close:
        return f"❌ Section '{sec}' not found."

    sec=close[0]
    idxs=_section_to_indices[sec]

    matches=[]
    for i in idxs:
        nm=data.iloc[i]["Name_lower"]
        if name in nm or get_close_matches(name,[nm],n=1,cutoff=0.6):
            matches.append(i)

    if not matches:
        return f"❌ No student '{name}' found in '{sec.upper()}'"

    return "\n\n".join(_format(i) for i in matches)


def _ph_search(d):
    out=[]
    for key,idxs in _phone_to_indices.items():
        if d in key:
            out+=idxs
    if not out:
        return f"❌ No student found for digits {d}"
    return "\n\n".join(_format(i) for i in out)

def _reg_search(r):
    if r in _reg_to_index:
        return _format(_reg_to_index[r])

    mask=data["Reg_clean"].str.contains(r,na=False,case=False)
    if mask.any():
        return "\n\n".join(_format(i) for i in data[mask].index)

    digits=re.sub(r"\D","",r)
    if len(digits)>=3:
        mask2=data["Reg_clean"].str.contains(digits,na=False)
        if mask2.any():
            return "\n\n".join(_format(i) for i in data[mask2].index)

    return f"❌ No reg.no matching '{r}'"

def _parent(name):
    idxs=_resolve(name)
    if not idxs: return f"❌ No student '{name}'."
    return "\n".join(f"📞 Parent Contact of {data.iloc[i]['Name_clean']}: {data.iloc[i]['Parent_contact_clean']}" for i in idxs)

def _student_contact(name):
    idxs=_resolve(name)
    if not idxs: return f"❌ No student '{name}'."
    return "\n".join(f"📱 Student Contact of {data.iloc[i]['Name_clean']}: {data.iloc[i]['Student_contact_clean']}" for i in idxs)

def _email(name):
    idxs=_resolve(name)
    if not idxs: return f"❌ No student '{name}'."
    return "\n".join(f"✉️ Email of {data.iloc[i]['Name_clean']}: {data.iloc[i]['Email_clean']}" for i in idxs)

def _full_info(name):
    idxs=_resolve(name)
    if not idxs: return f"❌ No student '{name}'."
    return "\n\n".join(_format(i) for i in idxs)

def _name(name):
    idxs=_resolve(name)
    if not idxs:
        return f"❌ No student '{name}'."

    if len(idxs)==1:
        return _format(idxs[0])

    return "⚠️ Multiple students found:\n\n" + "\n\n".join(_format(i) for i in idxs)


# ==============================
# MAIN API
# ==============================
def get_student_info(query: str) -> str:
    if not query:
        return "❌ Please enter a query."

    intent=_detect_intent(query)
    t=intent["intent"]
    v=intent.get("value","")

    if t=="help": return HELP_TEXT
    if t=="name_in_section": return _name_in_section(v[0],v[1])
    if t=="list_section": return _section_list(v)
    if t=="search_by_phone": return _ph_search(v)
    if t=="search_by_reg": return _reg_search(v)
    if t=="fetch_parent": return _parent(v)
    if t=="fetch_student_contact": return _student_contact(v)
    if t=="fetch_email": return _email(v)
    if t=="fetch_full_info": return _full_info(v)
    if t=="search_by_name": return _name(v)

    return "❌ Unknown error."
