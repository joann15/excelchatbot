import os
import json
from flask import Flask, request, jsonify, send_file
from dotenv import load_dotenv
from openai import OpenAI
from flask_cors import CORS
from google_drive import (
    upload_file,
    download_file,
    update_file
)
import requests
from openpyxl.styles import PatternFill
from datetime import datetime

load_dotenv(override=True)

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
key = os.getenv("OPENAI_API_KEY")
#print("Last 10:", key[-10:])

app = Flask(__name__)
CORS(app)

# ---------------------------
# In-memory DB
# ---------------------------
DB = []
LAST_UPLOADED_FILE = None
import sqlite3

def init_employee_db():
    conn = sqlite3.connect("employees.db")
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS employees (
            employee_name TEXT PRIMARY KEY,
            email TEXT NOT NULL
        )
    """)

    conn.commit()
    conn.close()
init_employee_db()

def get_employee_email(employee):

    conn = sqlite3.connect("employees.db")
    cursor = conn.cursor()

    first_name = employee.split()[0]

    cursor.execute(
        """
        SELECT email 
        FROM employees
        WHERE employee_name=?
        """,
        (first_name,)
    )

    row = cursor.fetchone()

    conn.close()

    if row:
        return row[0]

    return ""

def add_employee_db(name, email):
    conn = sqlite3.connect("employees.db")
    cursor = conn.cursor()

    cursor.execute(
        """
        INSERT OR REPLACE INTO employees(employee_name,email)
        VALUES(?,?)
        """,
        (name, email)
    )

    conn.commit()
    conn.close()

def update_employee(name, email):

    conn = sqlite3.connect("employees.db")
    cursor = conn.cursor()

    cursor.execute("""
        UPDATE employees
        SET email=?
        WHERE LOWER(employee_name)=LOWER(?)
    """, (email, name))

    conn.commit()
    conn.close()

def delete_employee(name):

    conn = sqlite3.connect("employees.db")
    cursor = conn.cursor()

    cursor.execute("""
        DELETE FROM employees
        WHERE LOWER(employee_name)=LOWER(?)
    """, (name,))

    conn.commit()
    conn.close()


# ---------------------------
# 4. Extract tasks
# ---------------------------
from openpyxl import load_workbook

def extract_tasks(file):

    wb = load_workbook(file)
    ws = wb.active

    headers = [cell.value for cell in ws[1]]

    open_col = headers.index("Open")
    close_col = headers.index("Close")

    status_map = {
        "FF6AA84F": "Done",
        "FFB7B7B7": "Due",
        "FF999999": "Due",
        "FF6FA8DC": "Half-Done",
        "FFFF0000": "Redo",
        "FFE84499": "Late",
        "FFBF8E00": "On Hold",
        "FF00FFFF": "Almost Ready",
        "FF7A3F00": "Just Started",
        "FFFFF2CC": "NA"
    }

    tasks = []
    
    for row in ws.iter_rows(min_row=2):
       # print([cell.value for cell in row])

        task = row[1].value
        
        employee_cells = [
            row[i].value
            for i in range(2, open_col)
            ]
        if all(v is None for v in employee_cells):
            continue

        # Skip blank rows
        if not task:
            continue
    
        
        open_date = row[open_col].value
        close_date = row[close_col].value
        
        for col_idx in range(2, open_col):
            employee = headers[col_idx]
            cell = row[col_idx]

            color = str(cell.fill.start_color.rgb)
            #print(f"Task={task}, Employee={employee}, Color={color}")

            print(task, employee, color)
            
            status = status_map.get(color)

            if status:
                tasks.append({
                    "task": task,
                    "employee": employee,
                    "status": status,
                    "color": "#" + color[-6:],   
                    "open": str(open_date),
                    "close": str(close_date)
                    })
                    
    print("Total tasks extracted:", len(tasks))
    
    if tasks:
        print("Last task:", tasks[-1])
                
    return tasks

N8N_WEBHOOK = "https://excelchatbot-n8n-production.up.railway.app/webhook/task-manager"


def send_to_n8n(
    action,
    task,
    employees,
    open_date,
    close_date,
    drive_file_id,
    field="",
    value="",
    updates=None
):
    
    emails = []

    for employee in employees:
        email = get_employee_email(employee)

        if email:
            emails.append(email)

    payload = {
        "action": action,
        "task": task,
        "employees": employees,
        "emails": emails,
        "open": open_date,
        "close": close_date,
        "field": field,
        "value": value,
        "updates": updates or [],
        "drive_file_id": drive_file_id
    }

    print("===== N8N PAYLOAD =====")
    print(payload)

    res = requests.post(N8N_WEBHOOK, json=payload)

    return res.status_code == 200

def find_task_details(excel_path, task_name):

    wb = load_workbook(excel_path)
    ws = wb.active

    headers = [cell.value for cell in ws[1]]

    open_col = headers.index("Open")

    status_colors = {
        "FF6AA84F",
        "FFB7B7B7",
        "FF999999",
        "FF6FA8DC",
        "FFFF0000",
        "FFE84499",
        "FFBF8E00",
        "FF00FFFF",
        "FF7A3F00",
        "FFFFF2CC"
    }


    employees = []


    for r in range(2, ws.max_row + 1):

        if ws.cell(r,2).value == task_name:

            for c in range(3, open_col+1):

                cell = ws.cell(r,c)

                color = str(cell.fill.start_color.rgb)

                if color in status_colors:
                    employees.append(headers[c-1])


            break
    emails = []    
    for emp in employees:
        email = get_employee_email(emp)

        print(emp, "->", email)

        if email:
            emails.append(email)

    return {
        "task": task_name,
        "employees": employees,
        "emails": emails
    }
    
@app.route("/")
def home():
    return "Job Card API is running!"

# ---------------------------
# 5. Upload endpoint
# ---------------------------
@app.route("/upload", methods=["POST"])
def upload():
    DB.clear()

    files = request.files.getlist("files")
    results = []
    
    import os
    import uuid
    
    os.makedirs("uploads", exist_ok=True)
    
    for file in files:
        
        unique_name = f"{uuid.uuid4()}_{file.filename}"
        filepath = os.path.join("uploads", unique_name)
        
        file.save(filepath)
        print(filepath)
        print(os.path.getsize(filepath))

        drive_file_id = upload_file(filepath)
        print("Uploaded to Google Drive:", drive_file_id)

        global LAST_UPLOADED_FILE
        global LAST_DRIVE_FILE_ID
        LAST_UPLOADED_FILE = filepath
        LAST_DRIVE_FILE_ID = drive_file_id

        tasks = extract_tasks(filepath)
        
        DB.append({
        "file": file.filename,
        "path": filepath,
        "tasks": tasks
    })

        #print(tasks[:10])


        results.append({
            "file": file.filename,
            "tasks_extracted": len(tasks)
        })

    return jsonify({
        "message": "Uploaded + normalized",
        "files": len(files),
        "results": results
    })

@app.route("/create-task", methods=["POST"])
def create_task():

    try:
        data = request.json

        print("===== CREATE TASK =====")
        print(data)

        task = data["task"]
        employees = data["employees"]
        open_date = data["open"]
        close_date = data["close"]
        drive_file_id = data["drive_file_id"]

        # Download Excel from Drive
        local_file = download_file(drive_file_id)

        wb = load_workbook(local_file)
        ws = wb.active

        headers = [cell.value for cell in ws[1]]

        print("Headers:", headers)

        # Find employee columns
        employee_cols = []

        for employee in employees:
            found = False

            for i, header in enumerate(headers):
                if header and employee.lower() in str(header).lower():
                    employee_cols.append(i + 1)
                    found = True
                    break

            if not found:
                return jsonify({
                    "error": f"Employee '{employee}' not found"
                }), 400


        # Find empty row
        new_row = ws.max_row + 1

        for r in range(2, ws.max_row + 1):
            if ws.cell(r, 2).value is None:
                new_row = r
                break


        # Task column
        task_col = 2

        open_col = headers.index("Open") + 1
        close_col = headers.index("Close") + 1


        # Insert task
        ws.cell(new_row, task_col).value = task


        open_dt = datetime.strptime(open_date, "%Y-%m-%d")
        close_dt = datetime.strptime(close_date, "%Y-%m-%d")


        ws.cell(new_row, open_col).value = open_dt
        ws.cell(new_row, close_col).value = close_dt


        ws.cell(new_row, open_col).number_format = "dd-mmm"
        ws.cell(new_row, close_col).number_format = "dd-mmm"


        # Assign employees
        for col in employee_cols:

            cell = ws.cell(new_row, col)

            cell.value = 1

            cell.fill = PatternFill(
                fill_type="solid",
                start_color="FF7A3F00",
                end_color="FF7A3F00"
            )


        wb.save(local_file)


        # Upload updated Excel
        update_file(
            drive_file_id,
            local_file
        )


        # Refresh chatbot memory
        new_tasks = extract_tasks(local_file)

        if DB:
            DB[0]["tasks"] = new_tasks


        os.remove(local_file)


        print("CREATE SUCCESS")


        return jsonify({
            "success": True,
            "task": task,
            "employees": employees
        })


    except Exception as e:

        print("CREATE ERROR:", e)

        return jsonify({
            "error": str(e)
        }),500
    
# update excel
from openpyxl import load_workbook

@app.route("/update-excel", methods=["POST"])
def update_excel():



    data = request.json

    print("Received:", data)

    task = data["task"]
    employees = data["employees"]
    open_date = data["open"]
    close_date = data["close"]
    drive_file_id = data["drive_file_id"]
    local_file = download_file(drive_file_id)
    wb = load_workbook(local_file)
    ws = wb.active

    headers = [cell.value for cell in ws[1]]

    print(headers)

    employee_cols = []
    
    for employee in employees:
        found = False
        
        for i, header in enumerate(headers):
            if header is not None and employee.lower() in str(header).lower():
                employee_cols.append(i + 1)
                found = True
                break
            
        if not found:
            return jsonify({
                "error": f"Employee '{employee}' not found."
                }), 400
    
    task_col = 2
    open_col = headers.index("Open") + 1
    close_col = headers.index("Close") + 1

    new_row = None
    
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 2).value is None:
            new_row = r
            break
        
    if new_row is None:
        new_row = ws.max_row + 1

    ws.cell(new_row, task_col).value = task

    open_dt = datetime.strptime(open_date, "%Y-%m-%d")
    close_dt = datetime.strptime(close_date, "%Y-%m-%d")
    
    ws.cell(new_row, open_col).value = open_dt
    ws.cell(new_row, close_col).value = close_dt
    
    ws.cell(new_row, open_col).number_format = "dd-mmm"
    ws.cell(new_row, close_col).number_format = "dd-mmm"

    for col in employee_cols:
        cell = ws.cell(new_row, col)
        
        cell.value = 1

        cell.fill = PatternFill(
            fill_type="solid",
            start_color="FF7A3F00",
            end_color="FF7A3F00")

    wb.save(local_file)
    update_file(drive_file_id, local_file)

    
    new_tasks = extract_tasks(local_file)
    
    print("New tasks extracted:", len(new_tasks))
    for t in new_tasks[-5:]:
        print(t)
    
    print("Downloaded file:", local_file)

    if DB:
        DB[0]["tasks"] = new_tasks

        print("Excel updated!")

        try:
            os.remove(local_file)
        except Exception as e:
            print("Couldn't delete temporary file:", e)

        return jsonify({
            "success": True,
            "task": task,
            "employees": employees,
            "emails": data.get("emails", []),
            "open": open_date,
            "close": close_date
            })


# ---------------------------
# 6. Dashboard API
# ---------------------------
@app.route("/dashboard", methods=["GET"])
def dashboard():

    employee_counts = {}
    employee_status = {}
    unique_tasks = set()
    status_counts = {}
    status_colors = {}
    
    if DB:
        wb = load_workbook(DB[0]["path"])
        ws = wb.active

        headers = [cell.value for cell in ws[1]]

        open_col = headers.index("Open")
        
        for employee in headers[2:open_col]:
            if employee:
                employee_counts.setdefault(employee, 0)
                employee_status.setdefault(employee, {})

    print("===== DASHBOARD =====")
            
    for doc in DB:
        print("Dashboard:", doc["file"], len(doc["tasks"]))

    for doc in DB:
        for t in doc["tasks"]:

            # Unique tasks
            unique_tasks.add(t["task"])

            emp = t["employee"]
            status = t["status"]
            status_colors[status] = t["color"]

            # Total tasks per employee
            employee_counts[emp] = employee_counts.get(emp, 0) + 1

            # Overall status counts
            status_counts[status] = status_counts.get(status, 0) + 1

            # Status counts per employee
            if emp not in employee_status:
                employee_status[emp] = {}

            employee_status[emp][status] = (
                employee_status[emp].get(status, 0) + 1
            )

    return jsonify({
    "total_tasks": len(unique_tasks),
    "total_employees": len(employee_counts),
    "tasks_per_employee": employee_counts,
    "status_breakdown": status_counts,
    "status_colors": status_colors,
    "employee_status": employee_status
})
# ---------------------------
# 7. Chat API
# ---------------------------
from collections import defaultdict

@app.route("/chat", methods=["POST"])
def chat():

    # -------------------------
    # STEP 1 - Get employee message
    # -------------------------
    query = request.json.get("message", "")

    # -------------------------
    # STEP 2 - Let GPT determine
    # what action the employee wants
    # -------------------------
    command = client.chat.completions.create(
        model="gpt-4.1",
        response_format={"type": "json_object"},
        messages=[
            {
                "role": "system",
                "content": """
You are a command detector.

Always respond with a valid JSON object.
Do not return markdown.
Do not return explanations.
Your response must always be JSON.

If the user wants to CREATE a task return:

{
    "action":"create",
    "task":"",
    "employees":[],
    "open":"",
    "close":""
}

If the user wants to DELETE a task return:

{
    "action":"delete",
    "task":""
}

If the user wants to update ONE field:

{
  "action":"update",
  "task":"LCM Testing",
  "employee":"Sneha",
  "updates":[
    {
      "field":"status",
      "value":"Done"
    }
  ]
}

If the user wants to ADD an employee return:

{
  "action":"add_employee",
  "employee":"",
  "email":""
}

If the user wants to update an employee email, return:

{
    "action":"update_employee",
    "name":"",
    "email":""
}

If the user wants to delete an employee, return:

{
    "action":"delete_employee",
    "name":""
}

If the user wants to update MULTIPLE fields:
{
  "action":"update",
  "task":"LCM Testing",
  "employee":"",
  "updates":[
    {
      "field":"status",
      "value":"Done"
    },
    {
      "field":"closeDate",
      "value":"2026-08-30"
    }
  ]
}

For status updates:
- Always include the employee whose status should change.
- If no employee is mentioned, set employee to "".

Otherwise return:

{
    "action":"chat"
}
"""
            },
            {
                "role": "user",
                "content": query
            }
        ]
    )

    command_json = json.loads(
        command.choices[0].message.content
    )

    print("COMMAND:", command_json)

    # -------------------------
    # Make sure a Job Card exists
    # -------------------------
    if command_json["action"] != "chat":

        if len(DB) == 0:
            return jsonify({
                "answer": "Please upload a Job Card first."
            })

        excel_path = DB[0]["path"]

    # ====================================================
    # CREATE TASK
    # ====================================================
    if command_json["action"] == "create":

        success = send_to_n8n(
            "create",
            command_json.get("task", ""),
            command_json.get("employees", []),
            command_json.get("open", ""),
            command_json.get("close", ""),
            LAST_DRIVE_FILE_ID
        )

        if success:
            return jsonify({
                "answer":
                f"Task '{command_json['task']}' assigned to "
                f"{', '.join(command_json['employees'])}."
            })

        return jsonify({
            "answer": "Failed to create task."
        })

    # ====================================================
    # DELETE TASK
    # ====================================================
    elif command_json["action"] == "delete":

        local_file = download_file(LAST_DRIVE_FILE_ID)
        details = find_task_details(
            local_file,
            command_json["task"]
        )
        
        os.remove(local_file)

        success = send_to_n8n(
            "delete",
            command_json["task"],
            details["employees"],
            "",
            "",
            LAST_DRIVE_FILE_ID
            )

        if success:
            return jsonify({
                "answer": f"Task '{command_json['task']}' deleted."
            })

        return jsonify({
            "answer": "Failed to delete task."
        })
    
    # ====================================================
    # UPDATE TASK
    # ====================================================
    elif command_json["action"] == "update":
        task = command_json["task"]
        updates = command_json.get("updates", [])
        
        if command_json.get("employee"):
            employees = [command_json["employee"]]
        else:
            local_file = download_file(LAST_DRIVE_FILE_ID)
            details = find_task_details(
            local_file,
            task
        )
            os.remove(local_file)
            employees = details["employees"]
        
        success = True
        performed_updates = []
        
        for update in updates:
            response = requests.post(
            "https://excelchatbot.onrender.com/update-task",
            json={
                "task": task,
                "field": update["field"],
                "value": update["value"],
                "employees": employees,
                "selected_employee": command_json.get("employee", ""),
                "drive_file_id": LAST_DRIVE_FILE_ID
                }
                )
        
            if response.status_code == 200:
                performed_updates.append({
                "field": update["field"],
                "value": update["value"]
                })
            else:
                success = False
            
        if success:
            send_to_n8n(
                action="update",
                task=task,
                employees=employees,
                open_date="",
                close_date="",
                drive_file_id=LAST_DRIVE_FILE_ID,
                updates=performed_updates
                )
            
            return jsonify({
            "answer": f"Task '{task}' updated successfully and employee notified."
            })
        return jsonify({
                "answer": "Failed to update task."
                })
        
        # ====================================================
        # ADD EMPLOYEE
        # ====================================================
            
    elif command_json["action"] == "add_employee":
        
        add_employee_db(
            command_json["employee"],
            command_json["email"]
            )
        
        response = requests.post(
            "https://excelchatbot.onrender.com/add-employee",
            json={
                "employee": command_json["employee"],
                "email": command_json["email"],
                "drive_file_id": LAST_DRIVE_FILE_ID
                }
        )

        if response.status_code == 200:
            send_to_n8n(
                action="welcome_employee",
                task="",
                employees=[command_json["employee"]],
                open_date="",
                close_date="",
                drive_file_id=LAST_DRIVE_FILE_ID
                )
        
    
            return jsonify({
                "answer": f"{command_json['employee']} added successfully."
                })
            
        return jsonify({
            "answer": "Failed to add employee."
            })

    # ====================================================
    # NORMAL CHAT
    # ====================================================
    
    context = []

    emp_tasks = defaultdict(dict)

    for doc in DB:
        for t in doc["tasks"]:

            context.append(t)

            emp = t["employee"]
            task = t["task"]

            emp_tasks[emp][task] = {
                "status": t["status"],
                "open": t["open"],
                "close": t["close"]
            }

    res = client.chat.completions.create(
        model="gpt-4.1",
        messages=[
            {
                "role": "system",
                "content": """
You are an AI Project Assistant.

You answer ONLY from the provided job card data.

RULES

1. Be concise.
2. Never invent information.
3. Use bullet points when appropriate.
4. If nothing matches say:
"No matching tasks were found."
"""
            },
            {
                "role": "user",
                "content": f"""
JOB CARD DATA

{json.dumps(context, indent=2)}

QUESTION

{query}
"""
            }
        ]
    )

    return jsonify({
        "answer": res.choices[0].message.content
    })
def get_employee_email(employee):

    if not employee:
        return ""

    first_name = employee.split()[0]

    conn = sqlite3.connect("employees.db")
    cursor = conn.cursor()

    cursor.execute(
        """
        SELECT email
        FROM employees
        WHERE LOWER(employee_name)=LOWER(?)
        """,
        (first_name,)
    )

    row = cursor.fetchone()

    print("DB SEARCH:", first_name, "RESULT:", row)

    conn.close()

    if row:
        return row[0]

    return ""

@app.route("/delete-task", methods=["POST"])
def delete_task():
    data = request.json

    task = data["task"]
    drive_file_id = data["drive_file_id"]

    local_file = download_file(drive_file_id)

    # Get assigned employees BEFORE deleting the task
    task_details = find_task_details(local_file, task)

    employees = task_details.get("employees", [])
    emails = [get_employee_email(emp) for emp in employees]

    wb = load_workbook(local_file)
    ws = wb.active

    deleted = False

    for r in range(2, ws.max_row + 1):
        if str(ws.cell(r, 2).value).strip().lower() == task.strip().lower():
            ws.delete_rows(r)
            deleted = True
            break

    if not deleted:
        os.remove(local_file)

        return jsonify({
            "success": False,
            "message": "Task not found."
        }), 404

    wb.save(local_file)
    update_file(drive_file_id, local_file)

    new_tasks = extract_tasks(local_file)

    if DB:
        DB[0]["tasks"] = new_tasks

    print(f"Deleted task: {task}")

    try:
        os.remove(local_file)
    except Exception as e:
        print("Couldn't delete temporary file:", e)

    return jsonify({
        "success": True,
        "message": f"{task} deleted.",
        "employees": employees,
        "emails": emails
    })

@app.route("/task-details", methods=["POST"])
def task_details():
    try:
        data = request.json
        print("TASK DETAILS:", data)

        drive_file_id = data["drive_file_id"]

        local_file = download_file(drive_file_id)

        result = find_task_details(local_file, data["task"])

        employee_names = result.get("employees", [])

        emails = [get_employee_email(emp) for emp in employee_names]

        result["emails"] = emails

        result["drive_file_id"] = drive_file_id

        os.remove(local_file)

        return jsonify(result)

    except Exception as e:
        print("TASK DETAILS ERROR:", e)
        return jsonify({"error": str(e)}), 500
    
@app.route("/update-task", methods=["POST"])
def update_task():
    
    data = request.json
    
    print("========== UPDATE ==========")
    print(data)
    
    task = data["task"]
    field = data["field"].lower()
    value = data["value"]
    employees = data.get("employees", [])
    employee = data.get("selected_employee", "")
    emails = data.get("emails", [])
   

    print("Task:", task)
    print("Field:", field)
    print("Value:", value)
    print("Employees:", employees)
    print("Employee selected:", employee)


    field_map = {
        "name": "task",
        "opendate": "open",
        "closedate": "close"
    }

    field = field_map.get(field, field)

    drive_file_id = data["drive_file_id"]
    local_file = download_file(drive_file_id)

    wb = load_workbook(local_file)
    ws = wb.active

    headers = [cell.value for cell in ws[1]]

    open_col = headers.index("Open") + 1
    close_col = headers.index("Close") + 1

    # -----------------------
    # Find task row
    # -----------------------

    task_row = None

    for r in range(2, ws.max_row + 1):
        if str(ws.cell(r, 2).value).strip().lower() == task.strip().lower():
            task_row = r
            break

    if task_row is None:
        return jsonify({
            "success": False,
            "message": "Task not found."
        }), 404

    # ==================================================
    # STATUS UPDATE
    # ==================================================

    if field == "status":

        status_colors = {
            "done": "FF6AA84F",
            "due": "FFB7B7B7",
            "half-done": "FF6FA8DC",
            "redo": "FFFF0000",
            "late": "FFE84499",
            "on hold": "FFBF8E00",
            "almost ready": "FF00FFFF",
            "just started": "FF7A3F00",
            "na": "FFFFF2CC"
        }

        if value.lower() not in status_colors:
            return jsonify({
                "success": False,
                "message": "Unknown status."
            }), 400

        fill = PatternFill(
            fill_type="solid",
            start_color=status_colors[value.lower()],
            end_color=status_colors[value.lower()]
        )

        # ----------------------------
        # If employee specified
        # ----------------------------

        if employee:

            employee_col = None

            for i, header in enumerate(headers):
                print(i + 1, header)
                
                if header is None:
                    continue
                
                if employee.lower() in str(header).lower():
                    employee_col = i + 1
                    
                    print("Matched employee column:", employee_col)
                    break

            if employee_col is None:
                return jsonify({
                    "success": False,
                    "message": "Employee not found."
                }), 404

            cell = ws.cell(task_row, employee_col)

            if cell.value is not None:
                print("Old color:", cell.fill.start_color.rgb)
                print("Updating cell:", task_row, employee_col)
                cell.fill = fill
                print("New color:", cell.fill.start_color.rgb)

        # ----------------------------
        # Otherwise update everyone
        # ----------------------------

        else:

            for c in range(3, open_col):

                cell = ws.cell(task_row, c)

                if cell.value is not None:
                    cell.fill = fill

    # ==================================================
    # TASK NAME
    # ==================================================

    elif field == "task":

        ws.cell(task_row, 2).value = value

    # ==================================================
    # OPEN DATE
    # ==================================================

    elif field == "open":

        dt = datetime.strptime(value, "%Y-%m-%d")

        ws.cell(task_row, open_col).value = dt
        ws.cell(task_row, open_col).number_format = "dd-mmm"

    # ==================================================
    # CLOSE DATE
    # ==================================================

    elif field == "close":

        dt = datetime.strptime(value, "%Y-%m-%d")

        ws.cell(task_row, close_col).value = dt
        ws.cell(task_row, close_col).number_format = "dd-mmm"

    else:

        return jsonify({
            "success": False,
            "message": "Unknown field."
        }), 400

    wb.save(local_file)
    update_file(drive_file_id, local_file)
    print("Workbook saved.")

    # Refresh dashboard data
    new_tasks = extract_tasks(local_file)

    if DB:
        DB[0]["tasks"] = new_tasks
        

    try:
        os.remove(local_file)
    except Exception as e:
        print("Couldn't delete temporary file:", e)

    print("Task updated!")

    return jsonify({
        "success": True,
        "task": task,
        "field": field,
        "value": value,
        "employees": employees,
        "emails": emails
    })
@app.route("/add-employee", methods=["POST"])
def add_employee():

    data = request.json

    print("========== ADD EMPLOYEE ==========")
    print(data)

    employee = data["employee"]
    email = data["email"]
    drive_file_id = data["drive_file_id"]

    local_file = download_file(drive_file_id)

    wb = load_workbook(local_file)
    ws = wb.active

    headers = [cell.value for cell in ws[1]]

    # Find the Open column
    open_col = headers.index("Open") + 1

    # Prevent duplicate employee
    if employee in headers:
        os.remove(local_file)
        return jsonify({
            "success": False,
            "message": "Employee already exists."
        })

    # Save email in database
    add_employee_db(employee, email)

    # Insert new employee column BEFORE Open
    ws.insert_cols(open_col)

    # Write employee name in header
    ws.cell(row=1, column=open_col).value = employee

    # Keep remaining cells blank
    for r in range(2, ws.max_row + 1):
        ws.cell(r, open_col).value = None

    wb.save(local_file)
    update_file(drive_file_id, local_file)

    # Refresh dashboard
    new_tasks = extract_tasks(local_file)

    if DB:
        DB[0]["tasks"] = new_tasks

    try:
        os.remove(local_file)
    except Exception as e:
        print("Couldn't delete temporary file:", e)

    return jsonify({
        "success": True,
        "message": f"{employee} added successfully."
    })
 
@app.route("/download")
def download():

    global LAST_UPLOADED_FILE

    if not LAST_UPLOADED_FILE:
        return jsonify({
            "error": "No uploaded file."
        }), 404

    local_file = download_file(LAST_DRIVE_FILE_ID)

    return send_file(
        local_file,
        as_attachment=True,
        download_name="Updated_JobCard.xlsx"
    )
# ---------------------------
# RUN
# ---------------------------
if __name__ == "__main__":
    app.run(debug=False)