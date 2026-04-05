import sys
sys.path.append('.')
from university_chatbot import generate_sql
print("Library:", generate_sql("Where is the library?"))
print("HOD:", generate_sql("Where is the hod office?"))
