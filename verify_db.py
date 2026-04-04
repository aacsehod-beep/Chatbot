import sqlite3
conn = sqlite3.connect('university.db')

print('=== FACULTY ===')
for r in conn.execute('SELECT name, phone, designation, department FROM faculty LIMIT 3'):
    print(' ', r)

print('\n=== STUDENTS ===')
for r in conn.execute('SELECT reg_no, name, section, student_phone FROM students LIMIT 3'):
    print(' ', r)

print('\n=== TIMETABLE CSE-3A Monday ===')
for r in conn.execute("SELECT section, day, hour, subject, teacher FROM timetable WHERE section='CSE-3A' LIMIT 5"):
    print(' ', r)

print('\n=== WORKLOAD ===')
for r in conn.execute("SELECT faculty, day, hour, subject_section FROM workload WHERE faculty LIKE '%Swathi%' LIMIT 4"):
    print(' ', r)

print('\n=== ATTENDANCE below 75 ===')
for r in conn.execute("SELECT name, subject, percentage FROM attendance WHERE section='CSE-2A' AND percentage < 75 AND week='week1' LIMIT 5"):
    print(' ', r)

print('\n=== SECTIONS in students ===')
for r in conn.execute('SELECT DISTINCT section FROM students ORDER BY section'):
    print(' ', r[0])

conn.close()
print('\nVerification complete.')
