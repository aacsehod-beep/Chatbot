from flask import Flask, render_template, request, jsonify
import pandas as pd

# Import your real modules
from modules.faculty import get_faculty_info
from modules.attendance import gui_get_attendance
from modules.student import get_student_info, data as student_data
from modules.timetable import gui_get_timetable
from modules.workload import gui_get_workload, FILE_PATH

app = Flask(__name__)

# Preload data
student_sections = student_data['Section'].unique()
faculty_names = pd.ExcelFile(FILE_PATH).sheet_names

# ---------------- Query Processing ----------------
def process_query(user_input, feature=None):
    if feature == "Student":
        info = get_student_info(user_input)
        return f"<b>Student Info:</b><br>{info.replace(chr(10), '<br>')}"
    
    elif feature == "Faculty":
        info = get_faculty_info(user_input)
        return f"<b>Faculty Info:</b><br>{info.replace(chr(10), '<br>')}"
    
    elif feature == "Attendance":
        info = gui_get_attendance(user_input)
        # attendance returns string with newlines
        return f"<b>Attendance:</b><br>{info.replace(chr(10), '<br>')}"
    
    elif feature == "Timetable":
        timetable_lines = gui_get_timetable(user_input)  # returns list of strings
        html = "<b>Timetable:</b><br><ul>"
        for line in timetable_lines:
            if line.strip() == "":
                continue
            # Bold headers (days or timetable title)
            if line.startswith("---") or "Timetable" in line:
                html += f"<li><b>{line}</b></li>"
            else:
                html += f"<li>{line}</li>"
        html += "</ul>"
        return html
    
    elif feature == "Workload":
        info = gui_get_workload(user_input)
        return f"<b>Workload:</b><br>{info.replace(chr(10), '<br>')}"
    
    else:
        return (
            "🤖 I can answer questions about <b>Student, Faculty, Attendance, Timetable, or Workload</b>. "
            "Please specify clearly."
        )

# ---------------- Routes ----------------
@app.route("/")
def index():
    return render_template("index.html")

@app.route("/ask", methods=["POST"])
def ask():
    data = request.get_json()
    user_input = data.get("message", "")
    feature = data.get("feature", None)
    response = process_query(user_input, feature)
    return jsonify({"response": response})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
