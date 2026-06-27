from flask import Flask, request, send_file
import psycopg2
from psycopg2.extras import RealDictCursor
from openpyxl import Workbook
from datetime import datetime
import pytz
import os

app = Flask(__name__)

# ---------------------------------
# USERS
# ---------------------------------

USERS = {
    "Gaurav": "1234",
    "Arshi": "5678"
}

# ---------------------------------
# DATABASE URL
# ---------------------------------

DATABASE_URL = os.environ.get("DATABASE_URL")

if not DATABASE_URL:
    raise Exception("DATABASE_URL")

# ---------------------------------
# CONNECT DATABASE
# ---------------------------------

def get_connection():

    return psycopg2.connect(
        DATABASE_URL,
        cursor_factory=RealDictCursor
    )

# ---------------------------------
# CREATE TABLE
# ---------------------------------

def create_table():

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""

    CREATE TABLE IF NOT EXISTS attendance(

        id SERIAL PRIMARY KEY,

        name VARCHAR(100) NOT NULL,

        action VARCHAR(20) NOT NULL,

        attendance_date DATE NOT NULL,

        attendance_time TIME NOT NULL,

        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP

    )

    """)

    conn.commit()

    cur.close()

    conn.close()

create_table()

# ---------------------------------
# INDIA TIME
# ---------------------------------

india = pytz.timezone("Asia/Kolkata")
# ---------------------------------
# SAVE ATTENDANCE
# ---------------------------------

def save_data(name, action):

    conn = get_connection()
    cur = conn.cursor()

    try:

        now = datetime.now(india)

        attendance_date = now.date()

        attendance_time = now.time().replace(microsecond=0)

        # ---------------------------------
        # DUPLICATE CHECK
        # Same User + Same Action
        # Within Last 1 Minute
        # ---------------------------------

        cur.execute("""

        SELECT id

        FROM attendance

        WHERE

        name=%s

        AND action=%s

        AND created_at >= NOW() - INTERVAL '1 minute'

        LIMIT 1

        """,

        (name, action))

        duplicate = cur.fetchone()

        if duplicate:

            cur.close()

            conn.close()

            return False

        # ---------------------------------
        # INSERT RECORD
        # ---------------------------------

        cur.execute("""

        INSERT INTO attendance

        (

            name,

            action,

            attendance_date,

            attendance_time

        )

        VALUES

        (

            %s,

            %s,

            %s,

            %s

        )

        """,

        (

            name,

            action,

            attendance_date,

            attendance_time

        ))

        conn.commit()

        cur.close()

        conn.close()

        return True

    except Exception as e:

        conn.rollback()

        cur.close()

        conn.close()

        print(e)

        return False


# ---------------------------------
# VERIFY PIN
# ---------------------------------

def verify_pin(name, pin):

    if name not in USERS:

        return False

    if USERS[name] != pin:

        return False

    return True
    # ---------------------------------
# HOME PAGE
# ---------------------------------

@app.route("/")
def home():

    options = ""

    for user in USERS.keys():
        options += f'<option value="{user}">{user}</option>'

    return f"""

<!DOCTYPE html>

<html>

<head>

<meta charset="UTF-8">

<meta name="viewport" content="width=device-width, initial-scale=1.0">

<title>SGMEA Attendance</title>

<style>

*{{
margin:0;
padding:0;
box-sizing:border-box;
font-family:Arial,sans-serif;
}}

body{{
background:#f2f5f9;
display:flex;
justify-content:center;
align-items:center;
height:100vh;
}}

.card{{
background:white;
width:360px;
padding:25px;
border-radius:15px;
box-shadow:0 10px 25px rgba(0,0,0,.15);
}}

h2{{
text-align:center;
margin-bottom:20px;
color:#0b5ed7;
}}

select,
input{{
width:100%;
padding:12px;
margin-top:10px;
border-radius:10px;
border:1px solid #ccc;
font-size:16px;
}}

button{{
width:100%;
padding:13px;
margin-top:12px;
border:none;
border-radius:10px;
font-size:16px;
cursor:pointer;
}}

.in{{
background:#198754;
color:white;
}}

.out{{
background:#dc3545;
color:white;
}}

.download{{
background:#0d6efd;
color:white;
text-decoration:none;
display:block;
text-align:center;
padding:13px;
margin-top:12px;
border-radius:10px;
}}

.footer{{
margin-top:18px;
text-align:center;
font-size:13px;
color:gray;
}}

</style>

</head>

<body>

<div class="card">

<h2>SGMEA Attendance</h2>

<form action="/mark" method="POST">

<select name="name" required>

{options}

</select>

<input
type="password"
name="pin"
placeholder="Enter PIN"
required>

<button
class="in"
type="submit"
name="action"
value="Check In">

✅ Check In

</button>

<button
class="out"
type="submit"
name="action"
value="Check Out">

❌ Check Out

</button>

</form>

<a
class="download"
href="/download">

📥 Download Attendance

</a>

<div class="footer">

Powered by SGMEA

</div>

</div>

</body>

</html>

"""
    # ---------------------------------
# MARK ATTENDANCE
# ---------------------------------

@app.route("/mark", methods=["POST"])
def mark():

    name = request.form.get("name", "").strip()
    pin = request.form.get("pin", "").strip()
    action = request.form.get("action", "").strip()

    # -------------------------------
    # VALIDATION
    # -------------------------------

    if name == "" or pin == "" or action == "":

        return """
        <script>
        alert("Please fill all fields.");
        window.location.href="/";
        </script>
        """

    # -------------------------------
    # PIN CHECK
    # -------------------------------

    if not verify_pin(name, pin):

        return """
        <script>
        alert("Wrong PIN !");
        window.location.href="/";
        </script>
        """

    # -------------------------------
    # SAVE ATTENDANCE
    # -------------------------------

    saved = save_data(name, action)

    if saved:

        return f"""
        <script>

        alert("{name}\\n\\n{action} Successful.");

        window.location.href="/";

        </script>
        """

    else:

        return f"""
        <script>

        alert("{name}\\n\\nDuplicate attendance detected.\\nPlease wait one minute.");

        window.location.href="/";

        </script>
        """
        # ---------------------------------
# DOWNLOAD EXCEL
# ---------------------------------

@app.route("/download")
def download():

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""

    SELECT

        name,
        action,
        attendance_date,
        attendance_time

    FROM attendance

    ORDER BY created_at DESC

    """)

    records = cur.fetchall()

    cur.close()
    conn.close()

    wb = Workbook()
    ws = wb.active

    ws.title = "Attendance"

    # -------------------------
    # HEADER
    # -------------------------

    ws.append([

        "Sr No",

        "Member Name",

        "Attendance",

        "Date",

        "Time"

    ])

    # -------------------------
    # DATA
    # -------------------------

    sr = 1

    for row in records:

        ws.append([

            sr,

            row["name"],

            row["action"],

            row["attendance_date"].strftime("%d-%m-%Y"),

            row["attendance_time"].strftime("%I:%M:%S %p")

        ])

        sr += 1

    # -------------------------
    # AUTO WIDTH
    # -------------------------

    for column in ws.columns:

        length = max(

            len(str(cell.value))

            if cell.value else 0

            for cell in column

        )

        ws.column_dimensions[
            column[0].column_letter
        ].width = length + 5

    # -------------------------
    # SAVE FILE
    # -------------------------

    DOWNLOAD_FOLDER = os.path.join(os.getcwd(), "downloads")

    os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

    FILE_PATH = os.path.join(

        DOWNLOAD_FOLDER,

        "SGMEA_Attendance.xlsx"

    )

    wb.save(FILE_PATH)

    wb.close()

    return send_file(

        FILE_PATH,

        as_attachment=True,

        download_name="SGMEA_Attendance.xlsx"

    )
    # ---------------------------------
# DASHBOARD
# ---------------------------------

@app.route("/dashboard")
def dashboard():

    conn = get_connection()
    cur = conn.cursor()

    # Total Records
    cur.execute("SELECT COUNT(*) AS total FROM attendance")
    total = cur.fetchone()["total"]

    # Total Check In
    cur.execute("""
        SELECT COUNT(*) AS total
        FROM attendance
        WHERE action='Check In'
    """)
    checkin = cur.fetchone()["total"]

    # Total Check Out
    cur.execute("""
        SELECT COUNT(*) AS total
        FROM attendance
        WHERE action='Check Out'
    """)
    checkout = cur.fetchone()["total"]

    # Today's Attendance
    cur.execute("""
        SELECT COUNT(*) AS total
        FROM attendance
        WHERE attendance_date=CURRENT_DATE
    """)
    today = cur.fetchone()["total"]

    cur.close()
    conn.close()

    return f"""

    <html>

    <head>

    <title>SGMEA Dashboard</title>

    <style>

    body{{
        font-family:Arial;
        background:#f3f5f7;
        padding:30px;
    }}

    .box{{
        background:white;
        padding:20px;
        margin:15px;
        border-radius:15px;
        box-shadow:0 0 10px rgba(0,0,0,.15);
        font-size:22px;
    }}

    a{{
        text-decoration:none;
        font-size:18px;
    }}

    </style>

    </head>

    <body>

    <h2>SGMEA Attendance Dashboard</h2>

    <div class="box">

    📋 Total Records : <b>{total}</b>

    </div>

    <div class="box">

    ✅ Total Check In : <b>{checkin}</b>

    </div>

    <div class="box">

    ❌ Total Check Out : <b>{checkout}</b>

    </div>

    <div class="box">

    📅 Today's Attendance : <b>{today}</b>

    </div>

    <br>

    <a href="/">⬅ Back to Attendance</a>

    </body>

    </html>

    """


# ---------------------------------
# HEALTH CHECK
# ---------------------------------

@app.route("/health")
def health():

    return {

        "status":"online",

        "database":"connected"

    }


# ---------------------------------
# TOTAL RECORD API
# ---------------------------------

@app.route("/count")
def count():

    conn = get_connection()

    cur = conn.cursor()

    cur.execute("SELECT COUNT(*) AS total FROM attendance")

    total = cur.fetchone()["total"]

    cur.close()

    conn.close()

    return {

        "records":total

    }
    # ---------------------------------
# SEARCH MEMBER ATTENDANCE
# ---------------------------------

@app.route("/search", methods=["GET", "POST"])
def search():

    if request.method == "GET":

        options = ""

        for user in USERS.keys():
            options += f'<option value="{user}">{user}</option>'

        return f"""

        <html>

        <head>

        <title>Search Attendance</title>

        <style>

        body{{
        font-family:Arial;
        background:#f2f2f2;
        }}

        .box{{
        width:350px;
        background:white;
        margin:80px auto;
        padding:20px;
        border-radius:15px;
        box-shadow:0 0 10px rgba(0,0,0,.2);
        }}

        select,button{{
        width:100%;
        padding:12px;
        margin-top:10px;
        border-radius:10px;
        }}

        </style>

        </head>

        <body>

        <div class="box">

        <h2 align="center">Search Attendance</h2>

        <form method="POST">

        <select name="member">

        {options}

        </select>

        <button type="submit">

        Search

        </button>

        </form>

        </div>

        </body>

        </html>

        """

    member = request.form["member"]

    conn = get_connection()

    cur = conn.cursor()

    cur.execute("""

    SELECT

    action,

    attendance_date,

    attendance_time

    FROM attendance

    WHERE name=%s

    ORDER BY created_at DESC

    """,(member,))

    records = cur.fetchall()

    cur.close()

    conn.close()

    rows = ""

    sr = 1

    for r in records:

        rows += f"""

        <tr>

        <td>{sr}</td>

        <td>{r['action']}</td>

        <td>{r['attendance_date']}</td>

        <td>{r['attendance_time']}</td>

        </tr>

        """

        sr += 1

    return f"""

    <html>

    <head>

    <title>{member}</title>

    <style>

    body{{font-family:Arial;background:#f5f5f5;padding:20px;}}

    table{{border-collapse:collapse;width:100%;background:white;}}

    th,td{{border:1px solid #ddd;padding:10px;text-align:center;}}

    th{{background:#0d6efd;color:white;}}

    </style>

    </head>

    <body>

    <h2>{member} Attendance History</h2>

    <table>

    <tr>

    <th>Sr</th>

    <th>Action</th>

    <th>Date</th>

    <th>Time</th>

    </tr>

    {rows}

    </table>

    <br>

    <a href="/">⬅ Back</a>

    </body>

    </html>

    """
    # ---------------------------------
# ERROR HANDLER
# ---------------------------------

@app.errorhandler(404)
def page_not_found(e):

    return """

    <html>

    <head>

    <title>404</title>

    <style>

    body{

        font-family:Arial;

        background:#f5f5f5;

        text-align:center;

        padding-top:80px;

    }

    a{

        text-decoration:none;

        font-size:18px;

    }

    </style>

    </head>

    <body>

    <h1>404</h1>

    <h3>Page Not Found</h3>

    <br>

    <a href="/">🏠 Go Home</a>

    </body>

    </html>

    """,404


# ---------------------------------
# 500 ERROR
# ---------------------------------

@app.errorhandler(500)
def internal_error(e):

    return """

    <script>

    alert("Internal Server Error");

    window.location="/";

    </script>

    """,500


# ---------------------------------
# CLOSE DATABASE
# ---------------------------------

@app.teardown_appcontext
def close_connection(exception):

    pass


# ---------------------------------
# MAIN
# ---------------------------------

if __name__ == "__main__":

    port = int(os.environ.get("PORT",5000))

    app.run(

        host="0.0.0.0",

        port=port,

        debug=False

    )
