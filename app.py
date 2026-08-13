from flask import Flask, request, send_file, redirect, url_for, session
import psycopg2
from psycopg2.extras import RealDictCursor
from openpyxl import Workbook
from datetime import datetime
import pytz
import os

app = Flask(__name__)

# Secret Key for Session Encryption & Security
app.config['SECRET_KEY'] = os.environ.get("SECRET_KEY", "SGMEA_SUPER_SECURE_SGM_KEY_1441")

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

    # Seed Initial Users
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
    except Exception:
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
# HOME PORTAL
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

# ---------------------------------
# SECURED OWNER ADMIN PANEL
# ---------------------------------
@app.route("/owner", methods=["GET", "POST"])
def owner():
    message = ""
    authenticated = session.get("owner_auth", False)

    if request.method == "POST":
        action_type = request.form.get("action_type")

        # Step 1: Login Check with Master PIN
        if action_type == "login":
            entered_pin = request.form.get("owner_pin", "").strip()
            if entered_pin == "1441":
                session["owner_auth"] = True
                authenticated = True
            else:
                message = "<script>alert('Incorrect Master PIN! Access Denied.');</script>"

        # Step 2: Add New Member (Only when Authenticated)
        elif action_type == "add_member" and authenticated:
            new_name = request.form.get("new_name", "").strip()
            new_pin = request.form.get("new_pin", "").strip()
            role = request.form.get("role", "Candidate").strip()

            if new_name and new_pin:
                try:
                    conn = get_connection()
                    cur = conn.cursor()
                    cur.execute("INSERT INTO users (name, pin, role) VALUES (%s, %s, %s)", (new_name, new_pin, role))
                    conn.commit()
                    cur.close()
                    conn.close()
                    message = f"<script>alert('Member \"{new_name}\" Added Successfully!');</script>"
                except Exception:
                    message = "<script>alert('Error: User already exists!');</script>"

        # Logout
        elif action_type == "logout":
            session.pop("owner_auth", None)
            return redirect(url_for("owner"))

    # Render Screen Based on Auth Status
    if authenticated:
        users = get_users_map()
        user_rows = "".join([f"<tr><td style='padding:8px; border-bottom:1px solid #1e293b;'>{u}</td></tr>" for u in users.keys()])
        
        content = f"""
        {message}
        <p style="text-align:center; color:#34d399; font-weight:bold; margin-bottom:15px;">🔓 Owner Access Granted (Shubham Agrawal SGM)</p>
        <form method="POST">
        <input type="hidden" name="action_type" value="add_member">
        <input type="text" name="new_name" placeholder="Staff / Candidate Name" required autocomplete="off">
        <input type="password" name="new_pin" placeholder="Assign Secret PIN" required autocomplete="off">
        <select name="role">
            <option value="Employee">Staff Employee</option>
            <option value="Candidate">D.Pharm Candidate</option>
        </select>
        <button class="in" type="submit">➕ Register New Member</button>
        </form>

        <h4 style="margin-top:20px; color:#38bdf8; text-align:center;">Registered Members List</h4>
        <div style="max-height:140px; overflow-y:auto; margin-top:10px; background:#0f172a; padding:10px; border-radius:10px; border:1px solid #334155;">
        <table style="width:100%; color:white; font-size:14px;">{user_rows}</table>
        </div>

        <form method="POST" style="margin-top:10px;">
        <input type="hidden" name="action_type" value="logout">
        <button type="submit" style="background:#dc2626; color:white; padding:10px;">🔒 Lock Owner Panel</button>
        </form>
        """
    else:
        content = f"""
        {message}
        <p style="text-align:center; color:#94a3b8; font-size:13px; margin-bottom:15px;">Authentication Required</p>
        <form method="POST">
        <input type="hidden" name="action_type" value="login">
        <input type="password" name="owner_pin" placeholder="Enter Owner Master PIN" required autocomplete="off">
        <button class="in" type="submit">🔓 Unlock Admin Access</button>
        </form>
        """

    return f"""
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><title>Secured Owner Panel</title>{COMMON_STYLE}</head>
<body>
<div class="card">
<h3>👑 Owner Admin Control</h3>
{content}
<a href="/" class="download" style="background:#475569;">⬅ Back to Home Portal</a>
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

    if not name or not pin or not action:
        return """<script>alert("All fields are required."); window.location.href="/";</script>"""

    if not verify_pin(name, pin):
        return """<script>alert("Authentication Failed: Invalid PIN!"); window.location.href="/";</script>"""

    if save_data(name, action):
        return f"""<script>alert("{name}\\n\\n{action} Marked Successfully."); window.location.href="/";</script>"""
    else:
        return f"""<script>alert("Duplicate Entry Blocked! Please wait 1 minute."); window.location.href="/";</script>"""

# ---------------------------------
# REPORTS & DASHBOARD
# ---------------------------------
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

@app.route("/download_user", methods=["GET", "POST"])
def download_user():
    users = get_users_map()
    if request.method == "GET":
        options = "".join([f'<option value="{u}">{u}</option>' for u in users.keys()])
        return f"""
        <html><head><title>Export Individual Attendance</title>{COMMON_STYLE}</head>
        <body><div class="card"><h3>Export Individual Report</h3>
        <form method="POST"><select name="member">{options}</select>
        <button class="in" type="submit">📥 Download Excel Report</button></form>
        <a href="/" class="download" style="background:#475569;">⬅ Back</a></div></body></html>
        """

    member = request.form.get("member")
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT action, attendance_date, attendance_time FROM attendance WHERE name=%s ORDER BY attendance_date ASC, attendance_time ASC", (member,))
    records = cur.fetchall()
    cur.close()
    conn.close()

    daily_data = {}
    for r in records:
        d = r["attendance_date"]
        if d not in daily_data:
            daily_data[d] = {"in": None, "out": None}
        if r["action"] == "Check In" and not daily_data[d]["in"]:
            daily_data[d]["in"] = r["attendance_time"]
        elif r["action"] == "Check Out":
            daily_data[d]["out"] = r["attendance_time"]

    wb = Workbook()
    ws = wb.active
    ws.title = f"{member}_Report"
    ws.append(["Sr No", "Member Name", "Date", "Day", "Check In Time", "Check Out Time", "Total Spent Time"])

    sr = 1
    for date_val, times in daily_data.items():
        day_str = date_val.strftime("%A")
        date_str = date_val.strftime("%d-%m-%Y")
        in_time_str = times["in"].strftime("%I:%M:%S %p") if times["in"] else "N/A"
        out_time_str = times["out"].strftime("%I:%M:%S %p") if times["out"] else "N/A"

        spent_str = "N/A"
        if times["in"] and times["out"]:
            t_in = datetime.combine(date_val, times["in"])
            t_out = datetime.combine(date_val, times["out"])
            if t_out > t_in:
                diff = t_out - t_in
                hours, remainder = divmod(diff.seconds, 3600)
                minutes, _ = divmod(remainder, 60)
                spent_str = f"{hours}h {minutes}m"

        ws.append([sr, member, date_str, day_str, in_time_str, out_time_str, spent_str])
        sr += 1

    DOWNLOAD_FOLDER = os.path.join(os.getcwd(), "downloads")
    os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)
    FILE_PATH = os.path.join(DOWNLOAD_FOLDER, f"{member}_Attendance_Report.xlsx")
    wb.save(FILE_PATH)
    wb.close()
    return send_file(FILE_PATH, as_attachment=True, download_name=f"{member}_Attendance_Report.xlsx")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
