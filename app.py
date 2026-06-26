from flask import Flask, request, send_file, redirect
import sqlite3
import os
from datetime import datetime
import pytz
from openpyxl import Workbook

app = Flask(__name__)

# -----------------------------
# STORAGE
# -----------------------------
BASE_DIR = os.path.join(os.getcwd(), "SGMEA")
os.makedirs(BASE_DIR, exist_ok=True)

DB_FILE = os.path.join(BASE_DIR, "attendance.db")
EXCEL_FILE = os.path.join(BASE_DIR, "Records.xlsx")

# -----------------------------
# USERS + PIN
# -----------------------------
USERS = {
    "Gaurav": "1234",
    "Arshi": "5678"
}

# -----------------------------
# DATABASE
# -----------------------------
def init_db():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS attendance(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT,
        action TEXT,
        date TEXT,
        time TEXT,
        timestamp TEXT
    )
    """)

    conn.commit()
    conn.close()

init_db()

# -----------------------------
# SAVE ATTENDANCE
# -----------------------------
def save_data(name, action):

    india = pytz.timezone("Asia/Kolkata")
    now = datetime.now(india)

    date = now.strftime("%d-%m-%Y")
    time = now.strftime("%I:%M:%S %p")
    timestamp = now.strftime("%Y-%m-%d %H:%M:%S")

    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()

    cur.execute("""
    INSERT INTO attendance(name,action,date,time,timestamp)
    VALUES(?,?,?,?,?)
    """,(name,action,date,time,timestamp))

    conn.commit()
    conn.close()

# -----------------------------
# HOME PAGE
# -----------------------------
@app.route("/")
def home():

    options = ""

    for user in USERS:
        options += f"<option>{user}</option>"

    return f"""
<html>

<head>

<title>SGMEA Attendance</title>

<style>

body{{
background:#f3f5f8;
font-family:Arial;
display:flex;
justify-content:center;
align-items:center;
height:100vh;
margin:0;
}}

.card{{
background:white;
padding:25px;
width:340px;
border-radius:15px;
box-shadow:0 0 15px rgba(0,0,0,.2);
}}

select,input,button{{
width:100%;
padding:12px;
margin-top:10px;
border-radius:10px;
border:none;
font-size:16px;
}}

.in{{
background:#28a745;
color:white;
}}

.out{{
background:#dc3545;
color:white;
}}

.download{{
background:#007bff;
color:white;
text-decoration:none;
display:block;
padding:12px;
margin-top:12px;
text-align:center;
border-radius:10px;
}}

</style>

</head>

<body>

<div class="card">

<h2 align=center>SGMEA Attendance</h2>

<form action="/mark" method="post">

<select name="name">

{options}

</select>

<input
type="password"
name="pin"
placeholder="Enter PIN"
required>

<button
class="in"
name="action"
value="Check In">
Check In
</button>

<button
class="out"
name="action"
value="Check Out">
Check Out
</button>

</form>

<a
class="download"
href="/download">
Download Excel
</a>

</div>

</body>

</html>
"""
    # -----------------------------
# MARK ATTENDANCE
# -----------------------------
@app.route("/mark", methods=["POST"])
def mark():

    name = request.form["name"]
    action = request.form["action"]
    pin = request.form["pin"]

    # PIN Verification
    if name not in USERS or USERS[name] != pin:
        return """
        <script>
        alert("❌ Wrong PIN!");
        window.location.href="/";
        </script>
        """

    # Save Attendance
    save_data(name, action)

    return """
    <script>
    alert("✅ Attendance Saved Successfully");
    window.location.href="/";
    </script>
    """


# -----------------------------
# DOWNLOAD EXCEL
# -----------------------------
@app.route("/download")
def download():

    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()

    cur.execute("""
    SELECT
    id,
    name,
    action,
    date,
    time
    FROM attendance
    ORDER BY id DESC
    """)

    rows = cur.fetchall()

    conn.close()

    wb = Workbook()
    ws = wb.active

    ws.title = "Attendance"

    ws.append([
        "Sr No",
        "Name",
        "Action",
        "Date",
        "Time"
    ])

    sr = 1

    for row in rows:

        ws.append([
            sr,
            row[1],
            row[2],
            row[3],
            row[4]
        ])

        sr += 1

    # Auto Width
    for column_cells in ws.columns:

        length = max(len(str(cell.value)) if cell.value else 0 for cell in column_cells)

        ws.column_dimensions[column_cells[0].column_letter].width = length + 5

    wb.save(EXCEL_FILE)

    return send_file(
        EXCEL_FILE,
        as_attachment=True,
        download_name="SGMEA_Attendance.xlsx"
    )
    # -----------------------------
# HEALTH CHECK (Render)
# -----------------------------
@app.route("/health")
def health():
    return "OK"


# -----------------------------
# RECORD COUNT (Optional)
# -----------------------------
@app.route("/count")
def count():

    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()

    cur.execute("SELECT COUNT(*) FROM attendance")
    total = cur.fetchone()[0]

    conn.close()

    return f"Total Attendance Records : {total}"


# -----------------------------
# RUN APP
# -----------------------------
if __name__ == "__main__":

    port = int(os.environ.get("PORT", 5000))

    app.run(
        host="0.0.0.0",
        port=port,
        debug=False
    )
