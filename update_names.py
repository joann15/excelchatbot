import sqlite3

conn = sqlite3.connect("employees.db")
cursor = conn.cursor()

updates = [
    ("Soumyadeep P", "Soumyadeep"),
    ("Shrawani M", "Shrawani"),
    ("Sneha B", "Sneha"),
    ("Rahi M", "Rahi"),
    ("Vanshika G", "Vanshika")
]

for new_name, old_name in updates:
    cursor.execute(
        """
        UPDATE employees
        SET employee_name = ?
        WHERE employee_name = ?
        """,
        (new_name, old_name)
    )

conn.commit()

cursor.execute("SELECT * FROM employees")

for row in cursor.fetchall():
    print(row)

conn.close()