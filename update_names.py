import sqlite3

conn = sqlite3.connect("employees.db")
cursor = conn.cursor()

updates = [
    ("Ankit O", "Ankit"),
    ("Pranav Y", "Pranav"),
    ("Akash P", "Akash"),
    ("Gaurav M", "Gaurav"),
    ("Aditya S", "Aditya"),
    ("Tanvi B", "Tanvi"),
    ("Chandrashekhar L", "Chandrashekhar"),
]

for new_name, old_name in updates:
    cursor.execute(
        """
        UPDATE employees
        SET employee_name=?
        WHERE employee_name=?
        """,
        (new_name, old_name)
    )

conn.commit()

cursor.execute("SELECT * FROM employees")
print(cursor.fetchall())

conn.close()

print("Done")