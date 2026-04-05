"""
university_chatbot.py — AuroMate
Text-to-SQL pipeline: user query → Gemini generates SQL → SQLite → Gemini formats answer
"""

from flask import Flask, render_template, request, jsonify
import sqlite3
import json
import uuid
import os
import re
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv
from google import genai

app = Flask(__name__)

load_dotenv()

# ─────────────────────────────────────────────
# DB + Conversation setup
# ─────────────────────────────────────────────
DB_PATH = "university.db"
CONVERSATIONS_FILE = Path("conversations.json")
if not CONVERSATIONS_FILE.exists():
    CONVERSATIONS_FILE.write_text(json.dumps({}))

# ─────────────────────────────────────────────
# Gemini setup
# ─────────────────────────────────────────────
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
GEMINI_MODEL_NAME = os.getenv("GEMINI_MODEL", "gemini-1.5-flash")
gemini_client = None
if GEMINI_API_KEY:
    try:
        gemini_client = genai.Client(api_key=GEMINI_API_KEY)
    except Exception:
        gemini_client = None

# ─────────────────────────────────────────────
# DB schema description sent to Gemini
# ─────────────────────────────────────────────
DB_SCHEMA = """
Database: university.db (SQLite)

Tables and columns:

locations
  id, room_name, block, floor, room_number, category, directions
  -- Stores building, room navigation, and directions.
  -- Example: room_name='Staff Room 1', block='A-Block', floor='1st Floor', category='Staff', directions='First corridor'

faculty
  id, sno, name, doj, designation, phone, cug, role, official_email, personal_email, department, office_location
  -- 40 rows. Stores faculty contact & designation info.
  -- office_location matches room_name in locations table.
  -- Example: name='Ms. K. Jayasri', phone='9100000661', designation='Asst.Prof.,CSE,SoE', department='School of Engineering', office_location='Staff Room 1'

students
  id, reg_no, name, student_phone, parent_phone, email, section
  -- 348 rows. One row per student.
  -- section values: AIML-2A, AIML-2B, CSE-2A, etc.
  -- reg_no pattern: like '231U1R1001'

timetable
  id, section, day, hour, subject, teacher, room, class_incharge
  -- 596 rows. One row per section/day/hour slot.
  -- hour values: H1, H2, H3, H4, H5, H6, H7, H8
  -- room column matches room_name in locations table.
  -- Example: section='CSE-3A', subject='Software Engineering(T)', teacher='Dr.B.Pannalal', room='Room 204'

workload
  id, faculty, day, hour, subject_section
  -- 496 rows. One row per faculty/day/hour.

attendance
  id, week, section, reg_no, name, subject, held, attended, percentage
"""

# ─────────────────────────────────────────────
# Conversation helpers
# ─────────────────────────────────────────────
def save_conversation(conv_id, user_msg, bot_msg, feature="General"):
    convs = json.loads(CONVERSATIONS_FILE.read_text())
    if conv_id not in convs:
        convs[conv_id] = {
            "messages": [],
            "created_at": datetime.now().isoformat(),
            "title": "New Conversation"
        }
    convs[conv_id]["messages"].append({
        "user": user_msg,
        "bot": bot_msg,
        "feature": feature,
        "timestamp": datetime.now().isoformat()
    })
    if len(convs[conv_id]["messages"]) == 1:
        convs[conv_id]["title"] = user_msg[:50]
    convs[conv_id]["messages"] = convs[conv_id]["messages"][-200:]
    CONVERSATIONS_FILE.write_text(json.dumps(convs, indent=2))

def get_conversations():
    return json.loads(CONVERSATIONS_FILE.read_text())


def get_recent_context(conv_id, max_turns=4):
    if not conv_id:
        return ""
    convs = get_conversations()
    messages = convs.get(conv_id, {}).get("messages", [])[-max_turns:]
    lines = []
    for m in messages:
        lines.append("USER: " + str(m.get("user", "")))
        lines.append("ASSISTANT: " + str(m.get("bot", "")).replace("<br>", " "))
    return "\n".join(lines)


# ─────────────────────────────────────────────
# DB query runner
# ─────────────────────────────────────────────
def run_sql(sql):
    """Execute a SELECT query. Returns (columns, rows). Only SELECT allowed."""
    sql_stripped = sql.strip().lstrip(";").strip()
    if not re.match(r'(?i)^\s*SELECT\b', sql_stripped):
        raise ValueError("Only SELECT queries are allowed.")
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        cursor = conn.execute(sql_stripped)
        rows = cursor.fetchall()
        columns = [d[0] for d in cursor.description] if cursor.description else []
        return columns, [dict(r) for r in rows]
    finally:
        conn.close()


def format_rows_as_text(columns, rows, max_rows=50):
    """Convert query results to compact readable text for Gemini."""
    if not rows:
        return "No records found."
    limited = rows[:max_rows]
    lines = []
    for r in limited:
        parts = [
            f"{c}: {r.get(c, '')}"
            for c in columns
            if r.get(c) is not None and str(r.get(c, '')).strip() not in ('', 'nan', 'None')
        ]
        if parts:
            lines.append("  • " + " | ".join(parts))
    suffix = f"\n  (showing {max_rows} of {len(rows)} total records)" if len(rows) > max_rows else ""
    return "\n".join(lines) + suffix


# ─────────────────────────────────────────────
# Language detection
# ─────────────────────────────────────────────
def _detect_language(text):
    for ch in text:
        cp = ord(ch)
        if 0x0C00 <= cp <= 0x0C7F:
            return "Telugu"
        if 0x0900 <= cp <= 0x097F:
            return "Hindi"
    return None


def _is_conversational(text):
    patterns = [
        r'^(hi|hello|hey|good morning|good afternoon|good evening|hii+|helo|sup)\b',
        r'^(thank|thanks|thx|ok|okay|got it|sure|bye|goodbye|see you)',
        r'^(who are you|what can you do|help me|what do you do|tell me about yourself)',
        r'^(how are you|how r u)',
    ]
    t = text.strip().lower()
    return any(re.match(p, t) for p in patterns)


# ─────────────────────────────────────────────
# Core Text-to-SQL pipeline
# ─────────────────────────────────────────────
def generate_sql(user_input, recent_context=""):
    """Ask Gemini to generate a safe SQLite SELECT query from natural language."""
    if not gemini_client:
        return None

    prompt = f"""You are a SQL expert for a university database. Convert the user's question into a SQLite SELECT query.

{DB_SCHEMA}

RULES:
1. Output ONLY the raw SQL query — no markdown, no explanation, no code fences.
2. Only use SELECT statements. Never use INSERT, UPDATE, DELETE, DROP.
3. Use TRIM() on text fields when filtering (data may have trailing spaces).
4. Use LIKE '%value%' for partial name/text matches instead of exact equality.
5. For attendance defaulters/shortage: WHERE percentage < 75
6. Limit results to 50 rows. Add LIMIT 50 if no LIMIT already.
7. Return NULL if the question cannot be answered from this database.
8. For timetable day filtering, use LIKE '%Monday%' to handle trailing spaces.
9. When user asks about a specific student by name, use LIKE on the name column.
10. To find a faculty's office location details, JOIN faculty and locations: SELECT f.name, l.* FROM faculty f JOIN locations l ON f.office_location = l.room_name WHERE f.name LIKE '%Smith%'
11. To find a class room location for a section/timetable, JOIN timetable and locations: SELECT t.*, l.block, l.floor, l.directions FROM timetable t JOIN locations l ON t.room = l.room_name

Recent conversation:
{recent_context or "None"}

User question: {user_input}

SQL:"""

    try:
        result = gemini_client.models.generate_content(
            model=GEMINI_MODEL_NAME,
            contents=prompt,
        )
        sql = (result.text or "").strip()
        if sql.startswith("```"):
            sql = sql.split("```")[1]
            if sql.lower().startswith("sql"):
                sql = sql[3:]
        sql = sql.strip().rstrip(";")
        if sql.upper() == "NULL" or not sql:
            return None
        return sql
    except Exception:
        return None


def format_with_gemini(user_input, sql, columns, rows, recent_context=""):
    """Ask Gemini to write a conversational answer from query results."""
    if not gemini_client:
        return None, []

    data_text = format_rows_as_text(columns, rows)
    lang = _detect_language(user_input)
    lang_instruction = f"Respond entirely in {lang}.\n" if lang else ""

    prompt = f"""You are AuroMate, a friendly AI university assistant.

The user asked: "{user_input}"

SQL query executed:
{sql}

Query results:
{data_text}

{lang_instruction}
Instructions:
- Present the data naturally and conversationally.
- Use bullet points or numbered lists where it helps readability.
- If no results: apologize briefly and suggest what to try (spelling, full name, specify section).
- If many rows: summarize the key insight first, then list details.
- NEVER invent data — only use what's in the query results above.
- Be concise and warm.

Recent conversation:
{recent_context or "None"}

After your answer, on a new line write exactly:
SUGGESTIONS: <follow-up question 1> | <follow-up question 2> | <follow-up question 3>
(Each suggestion max 7 words. Make them relevant to what was just discussed.)"""

    try:
        result = gemini_client.models.generate_content(
            model=GEMINI_MODEL_NAME,
            contents=prompt,
        )
        text = (result.text or "").strip()
        suggestions = []
        if "SUGGESTIONS:" in text:
            main_text, raw_sugg = text.split("SUGGESTIONS:", 1)
            suggestions = [s.strip() for s in raw_sugg.split("|") if s.strip()][:3]
        else:
            main_text = text
        return main_text.strip(), suggestions
    except Exception:
        return None, []


def handle_conversational(user_input, recent_context=""):
    """Handle greetings and small-talk."""
    if not gemini_client:
        return (
            "Hello! I'm AuroMate, your university assistant. "
            "Ask me about students, faculty, attendance, timetables, or workload.",
            ["Show attendance for CSE-2A", "What is Dr. Swathi's schedule?", "List AIML-3A students"]
        )
    prompt = (
        "You are AuroMate, a friendly AI academic assistant for Aurora University.\n"
        "You help with student info, faculty contacts, attendance, timetables and faculty workload.\n"
        f"Recent chat:\n{recent_context or 'None'}\n\n"
        f"User said: \"{user_input}\"\n"
        "Reply naturally and warmly in 1-2 sentences.\n"
        "Then on a new line write:\n"
        "SUGGESTIONS: <question1> | <question2> | <question3>\n"
        "Suggest 3 relevant things the user can ask AuroMate (max 7 words each)."
    )
    try:
        result = gemini_client.models.generate_content(model=GEMINI_MODEL_NAME, contents=prompt)
        text = (result.text or "").strip()
        suggestions = []
        if "SUGGESTIONS:" in text:
            main_text, raw_sugg = text.split("SUGGESTIONS:", 1)
            suggestions = [s.strip() for s in raw_sugg.split("|") if s.strip()][:3]
        else:
            main_text = text
        return main_text.strip(), suggestions
    except Exception:
        return "Hello! I'm AuroMate. How can I help you today?", []


def _self_correct_sql(bad_sql, error_msg, user_input):
    """Ask Gemini to fix a broken SQL query."""
    if not gemini_client:
        return None
    prompt = f"""The following SQLite query failed with error: {error_msg}

Bad SQL:
{bad_sql}

{DB_SCHEMA}

Fix the SQL query. Output ONLY the corrected raw SQL — no explanation, no markdown.
If it cannot be fixed, output: NULL

Original user question: {user_input}

Fixed SQL:"""
    try:
        result = gemini_client.models.generate_content(model=GEMINI_MODEL_NAME, contents=prompt)
        fixed = (result.text or "").strip().rstrip(";")
        if fixed.upper() == "NULL" or not fixed:
            return None
        if fixed.startswith("```"):
            fixed = fixed.split("```")[1]
            if fixed.lower().startswith("sql"):
                fixed = fixed[3:]
        return fixed.strip()
    except Exception:
        return None


def _detect_feature_from_sql(sql):
    """Infer module label from SQL table references."""
    sql_lower = sql.lower()
    if "attendance" in sql_lower:
        return "Attendance"
    if "workload" in sql_lower:
        return "Workload"
    if "timetable" in sql_lower:
        return "Timetable"
    if "faculty" in sql_lower:
        return "Faculty"
    if "student" in sql_lower:
        return "Student"
    return "General"


# ─────────────────────────────────────────────
# Main query processor
# ─────────────────────────────────────────────
def process_query(user_input, feature=None, conv_id=None):
    """
    Full pipeline:
      1. Conversational check → friendly chat
      2. Gemini generates SQL
      3. SQL runs on SQLite
      4. Gemini formats results conversationally
      5. Fallback: raw data if Gemini formatting fails
    Returns (response_html, ai_used, detected_feature, suggestions)
    """
    recent_context = get_recent_context(conv_id)

    # Step 1: conversational
    if _is_conversational(user_input):
        text, suggestions = handle_conversational(user_input, recent_context)
        return text.replace("\n", "<br>"), True, "General", suggestions

    # Step 2: generate SQL
    sql = generate_sql(user_input, recent_context)
    if not sql:
        fallback = (
            "I wasn't able to process that query. "
            "You can ask me about: student contacts, faculty info, class timetables, "
            "faculty workload, or student attendance."
        )
        return fallback, False, "General", []

    # Step 3: run SQL
    try:
        columns, rows = run_sql(sql)
    except Exception as e:
        corrected_sql = _self_correct_sql(sql, str(e), user_input)
        if corrected_sql:
            try:
                columns, rows = run_sql(corrected_sql)
                sql = corrected_sql
            except Exception:
                return "I had trouble querying the database. Could you rephrase your question?", False, "General", []
        else:
            return "I had trouble querying the database. Could you rephrase your question?", False, "General", []

    detected_feature = _detect_feature_from_sql(sql)

    # Step 4: format with Gemini
    response_text, suggestions = format_with_gemini(user_input, sql, columns, rows, recent_context)
    if response_text:
        return response_text.replace("\n", "<br>"), True, detected_feature, suggestions

    # Step 5: raw fallback
    raw_text = format_rows_as_text(columns, rows)
    if not raw_text or raw_text == "No records found.":
        return "No matching records found in the database.", False, detected_feature, []
    return raw_text.replace("\n", "<br>"), False, detected_feature, []

# ─────────────────────────────────────────────
# Routes
# ─────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/ask", methods=["POST"])
def ask():
    data = request.get_json()
    user_input = (data.get("message", "") or "").strip()
    if not user_input:
        return jsonify({"response": "Please enter a message.", "ai_used": False,
                        "detected_feature": "General", "suggestions": []})
    feature = data.get("feature", None)
    conv_id = data.get("convId", None)

    response, ai_used, detected_feature, suggestions = process_query(user_input, feature, conv_id)

    if conv_id:
        save_conversation(conv_id, user_input, response, detected_feature)

    return jsonify({
        "response": response,
        "ai_used": ai_used,
        "detected_feature": detected_feature,
        "suggestions": suggestions
    })


@app.route("/ai-status", methods=["GET"])
def ai_status():
    return jsonify({
        "configured": bool(gemini_client),
        "provider": "Gemini",
        "model": GEMINI_MODEL_NAME,
    })


@app.route("/new-conversation", methods=["POST"])
def new_conversation():
    conv_id = str(uuid.uuid4())
    convs = get_conversations()
    convs[conv_id] = {
        "messages": [],
        "created_at": datetime.now().isoformat(),
        "title": "New Conversation"
    }
    CONVERSATIONS_FILE.write_text(json.dumps(convs, indent=2))
    return jsonify({"convId": conv_id})


@app.route("/conversations", methods=["GET"])
def conversations_endpoint():
    convs = get_conversations()
    result = []
    for cid, cdata in convs.items():
        msgs = cdata.get("messages", [])
        result.append({
            "id": cid,
            "title": cdata.get("title", "Conversation"),
            "created_at": cdata.get("created_at", ""),
            "message_count": len(msgs),
            "last_message": msgs[-1].get("user", "") if msgs else ""
        })
    result.sort(key=lambda x: x.get("created_at", ""), reverse=True)
    return jsonify(result)


@app.route("/conversation/<conv_id>", methods=["GET"])
def get_conversation(conv_id):
    convs = get_conversations()
    return jsonify(convs.get(conv_id, {}))


@app.route("/export/<conv_id>", methods=["GET"])
def export_conversation(conv_id):
    convs = get_conversations()
    conv = convs.get(conv_id, {})
    if not conv:
        return jsonify({"error": "Conversation not found"}), 404
    lines = [
        "AuroMate Conversation Export",
        f"ID: {conv_id}",
        f"Date: {conv.get('created_at', '')}",
        "=" * 50, ""
    ]
    for msg in conv.get("messages", []):
        lines.append(f"You: {msg.get('user', '')}")
        bot = str(msg.get("bot", "")).replace("<br>", "\n").replace("<b>", "").replace("</b>", "")
        lines.append(f"AuroMate: {bot}")
        lines.append("")
    return "\n".join(lines), 200, {
        "Content-Type": "text/plain",
        "Content-Disposition": f"attachment; filename=conversation_{conv_id[:8]}.txt"
    }


if __name__ == "__main__":
    if not os.path.exists(DB_PATH):
        print(f"ERROR: {DB_PATH} not found. Run: python migrate_to_db.py")
    else:
        app.run(debug=True, host="0.0.0.0", port=5000)

