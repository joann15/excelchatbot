import sqlite3

conn = sqlite3.connect("employees.db")

cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS employees (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    employee_name TEXT UNIQUE NOT NULL,
    email TEXT NOT NULL
)
""")

employees = [
    ("Soumyadeep", "joann34074@gmail.com"),
    ("Akash", "joannmathews123@gmail.com"),
    ("Gaurav", "jmm4860@g.rit.edu"),
    ("Ankit", "joann34074@gmail.com"),
    ("Pranav", "joannmathews123@gmail.com"),
    ("Aditya", "jmm4860@g.rit.edu"),
    ("Shrawani", "joann34074@gmail.com"),
    ("Apeksha", "joannmathews123@gmail.com"),
    ("Sneha", "jmm4860@g.rit.edu"),
    ("Rahi", "joann34074@gmail.com"),
    ("Vanshika", "joannmathews123@gmail.com"),
    ("Ruhikesh", "jmm4860@g.rit.edu"),
    ("Nitish", "joann34074@gmail.com"),
    ("Tanvi", "joannmathews123@gmail.com"),
    ("Chandrashekhar", "jmm4860@g.rit.edu")
]

cursor.executemany("""
INSERT OR IGNORE INTO employees(employee_name,email)
VALUES(?,?)
""", employees)

conn.commit()

conn.close()

print("Database created!")