from flask import Flask, request, send_file, redirect, url_for
import psycopg2
from psycopg2.extras import RealDictCursor
from openpyxl import Workbook
from datetime import datetime
import pytz
import os

app = Flask(__name__)

# Secret key for security signing
app.config['SECRET_KEY'] = os.environ.get("SECRET_KEY", "SGMEA_SUPER_SECRET_KEY_1441")

# ---------------------------------
# SECURITY HEADERS
# ---------------------------------
@app.after_request
def apply_security_headers(response):
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['X-Frame-Options'] = 'DENY'
    response.headers['X-XSS-Protection'] = '1; mode=block'
    return response

# ---------------------------------
# DATABASE URL & CONNECTION
# ---------------------------------
DATABASE_URL = os.environ.get("DATABASE_URL")

if not DATABASE_URL:
    raise Exception("DATABASE_URL environment variable missing")

def get_connection():
    return psycopg2.connect(
        DATABASE_URL,
        cursor_factory=RealDictCursor,
        sslmode="require"
    )

# ---------------------------------
# INITIALIZE DATABASE SCHEMAS
# ---------------------------------
def init_db():
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
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS users(
        id SERIAL PRIMARY KEY,
        name VARCHAR(100) UNIQUE NOT NULL,
        pin VARCHAR(20) NOT NULL,
        role VARCHAR(20) DEFAULT 'Candidate'
    );
    """)

    # Seed Master Owner and Staff
    cur.execute("SELECT COUNT(*) as count FROM users;")
    if cur.fetchone()["count"] == 0:
        cur.execute("INSERT INTO users (name, pin, role) VALUES (%s, %s, %s)", ("Shubham Agrawal SGM", "1441", "Owner"))
        cur.execute("INSERT INTO users (name, pin, role) VALUES (%s, %s, %s)", ("Gaurav", "1234", "Employee"))
        cur.execute("INSERT INTO users (name, pin, role) VALUES (%s, %s, %s)", ("Arshi", "5678", "Employee"))

    conn.commit()
    cur.close()
    conn.close()

init_db()

def get_users_map():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT name, pin FROM users ORDER BY name ASC")
    rows = cur.fetchall()
    cur.close()
    conn.close()
    return {r["name"]: r["pin"] for r in rows}

india = pytz.timezone("Asia/Kolkata")

def save_data(name, action):
    conn = get_connection()
    cur = conn.cursor()
    try:
        now = datetime.now(india)
        attendance_date = now.date()
        attendance_time = now.time().replace(microsecond=0)

        # Anti-spam / Duplicate Check (1 minute interval limit)
        cur.execute("""
        SELECT id FROM attendance
        WHERE name=%s AND action=%s AND created_at >= NOW() - INTERVAL '1 minute'
        LIMIT 1
        """, (name, action))

        if cur.fetchone():
            cur.close()
            conn.close()
            return False

        cur.execute("""
        INSERT INTO attendance (name, action, attendance_date, attendance_time)
        VALUES (%s, %s, %s, %s)
        """, (name, action, attendance_date, attendance_time))

        conn.commit()
        cur.close()
        conn.close()
        return True
    except Exception as e:
        conn.rollback()
        cur.close()
        conn.close()
        return False

def verify_pin(name, pin):
    users = get_users_map()
    return users.get(name) == pin

# ---------------------------------
# UI STYLING
# ---------------------------------
COMMON_STYLE = """
<style>
* { margin: 0; padding: 0; box-sizing: border-box; font-family: 'Segoe UI', sans-serif; }
body { background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); color: #f8fafc; min-height: 100vh; display: flex; justify-content: center; align-items: center; padding: 20px; }
.card { background: rgba(30, 41, 59, 0.85); backdrop-filter: blur(16px); border: 1px solid rgba(255, 255, 255, 0.1); width: 100%; max-width: 420px; padding: 30px; border-radius: 20px; box-shadow: 0 20px 40px rgba(0,0,0,0.5); }
h2, h3 { text-align: center; margin-bottom: 20px; color: #38bdf8; }
select, input { width: 100%; padding: 14px; margin-top: 12px; border-radius: 12px; border: 1px solid #334155; background: #0f172a; color: #f8fafc; font-size: 15px; outline: none; }
button { width: 100%; padding: 14px; margin-top: 14px; border: none; border-radius: 12px; font-size: 16px; font-weight: bold; cursor: pointer; transition: 0.3s; }
.in { background: linear-gradient(135deg, #10b981 0%, #059669 100%); color: white; }
.out { background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%); color: white; }
.download { background: linear-gradient(135deg, #3b82f6 0%, #1d4ed8 100%); color: white; text-decoration: none; display: block; text-align: center; padding: 14px; margin-top: 12px; border-radius: 12px; font-weight: bold; }
.owner-btn { background: linear-gradient(135deg, #8b5cf6 0%, #6d28d9 100%); color: white; text-decoration: none; display: block; text-align: center; padding: 12px; margin-top: 12px; border-radius: 12px; font-weight: bold; }
.footer { margin-top: 22px; text-align: center; font-size: 12px; color: #94a3b8; }
</style>
"""

# ---------------------------------
# ROUTES
# ---------------------------------
@app.route("/")
def home():
    users = get_users_map()
    options = "".join([f'<option value="{u}">{u}</option>' for u in users.keys()])
    return f"""
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><title>SGMEA Secure Portal</title>{COMMON_STYLE}</head>
<body>
<div class="card">
<h2>🛡️ SGMEA Attendance</h2>
<form action="/mark" method="POST">
<select name="name" required><option value="" disabled selected>Select Staff / Candidate</option>{options}</select>
<input type="password" name="pin" placeholder="Enter Confidential PIN" required autocomplete="off">
<button class="in" type="submit" name="action" value="Check In">✅ Check In</button>
<button class="out" type="submit" name="action" value="Check Out">❌ Check Out</button>
</form>
<a class="download" href="/download">📥 Download Complete Excel</a>
<a class="download" style="background:#475569;" href="/download_user">📊 Individual Member Report</a>
<a class="owner-btn" href="/owner">👑 Owner Control (Shubham Agrawal SGM)</a>
<div class="footer">Powered by SGMEA Secure Platform</div>
</div>
</body>
</html>
"""

@app.route("/owner", methods=["GET", "POST"])
def owner():
    if request.method == "POST":
        owner_pin = request.form.get("owner_pin", "").strip()
        new_name = request.form.get("new_name", "").strip()
        new_pin = request.form.get("new_pin", "").strip()
        role = request.form.get("role", "Candidate").strip()

        # MASTER OWNER PIN SECURITY
        if owner_pin != "1441":
            return """<script>alert('Unauthorized Access! Incorrect Master PIN.'); window.location.href='/owner';</script>"""

        if new_name and new_pin:
            try:
                conn = get_connection()
                cur = conn.cursor()
                cur.execute("INSERT INTO users (name, pin, role) VALUES (%s, %s, %s)", (new_name, new_pin, role))
                conn.commit()
                cur.close()
                conn.close()
                return f"""<script>alert('Member "{new_name}" Added Successfully!'); window.location.href='/owner';</script>"""
            except Exception:
                return """<script>alert('Error: User already exists!'); window.location.href='/owner';</script>"""

    users = get_users_map()
    user_rows = "".join([f"<tr><td style='padding:8px;'>{u}</td></tr>" for u in users.keys()])

    return f"""
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><title>Owner Control Panel</title>{COMMON_STYLE}</head>
<body>
<div class="card">
<h3>👑 Owner Admin Panel</h3>
<p style="text-align:center; color:#94a3b8; font-size:13px; margin-bottom:15px;">Welcome Shubham Agrawal SGM</p>
<form method="POST">
<input type="password" name="owner_pin" placeholder="Enter Owner Master PIN (1441)" required>
<input type="text" name="new_name" placeholder="Staff / Candidate Name" required>
<input type="text" name="new_pin" placeholder="Assign Secret PIN" required>
<select name="role">
    <option value="Employee">Staff Employee</option>
    <option value="Candidate">D.Pharm Candidate</option>
</select>
<button class="in" type="submit">➕ Register New Member</button>
</form>
<h4 style="margin-top:20px; color:#38bdf8; text-align:center;">Registered Members</h4>
<div style="max-height:140px; overflow-y:auto; margin-top:10px; background:#0f172a; padding:10px; border-radius:10px;">
<table style="width:100%; color:white; font-size:14px;">{user_rows}</table>
</div>
<a href="/" class="download" style="background:#475569;">⬅ Back to Portal</a>
</div>
</body>
</html>
"""

@app.route("/mark", methods=["POST"])
def mark():
    name = request.form.get("name", "").strip()
    pin = request.form.get("pin", "").strip()
    action = request.form.get("action", "").strip()

    if not name or not pin or not action:
        return """<script>alert("All fields are required."); window.location.href="/";</script>"""

    if not verify_pin(name, pin):
        return """<script>alert("Authentication Failed: Invalid PIN!"); window.location.href="/";</script>"""

    if save_data(name, action):
        return f"""<script>alert("{name}\\n\\n{action} Marked Successfully."); window.location.href="/";</script>"""
    else:
        return f"""<script>alert("Duplicate Entry Blocked! Please wait 1 minute."); window.location.href="/";</script>"""

@app.route("/download")
def download():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT name, action, attendance_date, attendance_time FROM attendance ORDER BY created_at DESC")
    records = cur.fetchall()
    cur.close()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Attendance"
    ws.append(["Sr No", "Member Name", "Attendance", "Date", "Time"])

    for sr, row in enumerate(records, 1):
        ws.append([sr, row["name"], row["action"], row["attendance_date"].strftime("%d-%m-%Y"), row["attendance_time"].strftime("%I:%M:%S %p")])

    DOWNLOAD_FOLDER = os.path.join(os.getcwd(), "downloads")
    os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)
    FILE_PATH = os.path.join(DOWNLOAD_FOLDER, "SGMEA_Attendance.xlsx")
    wb.save(FILE_PATH)
    wb.close()
    return send_file(FILE_PATH, as_attachment=True, download_name="SGMEA_Attendance.xlsx")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
