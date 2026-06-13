from flask import Flask, request
from openpyxl import Workbook, load_workbook
import pandas as pd
from datetime import datetime
import pytz

app = Flask(__name__)

from flask import send_file

@app.route("/download")
def download():
    return send_file("SGMEA/Records.xlsx", as_attachment=True)
# 🔥 STORAGE PATH (Termux / Android)
import os

BASE_DIR = os.path.join(os.getcwd(), "SGMEA")
os.makedirs(BASE_DIR, exist_ok=True)

FILE = os.path.join(BASE_DIR, "Records.xlsx")

# 🔐 USERS + PIN
USERS = {
    "Gaurav": "1234",
    "Arshi": "5678"
}

# 📊 CREATE FILE IF NOT EXISTS
if not os.path.exists(FILE):
    wb = Workbook()
    ws = wb.active
    ws.append(["Name", "Action", "Time"])
    wb.save(FILE)


# 💾 SAVE FUNCTION (NO OVERWRITE, ONLY APPEND)
def save_data(name, action):
    wb = load_workbook(FILE)
    ws = wb.active

    india = pytz.timezone("Asia/Kolkata")
    current_time = datetime.now(india).strftime("%Y-%m-%d %H:%M:%S")

    new_data = pd.DataFrame([[name, current_time]], columns=["Name", "Time"])

    india = pytz.timezone("Asia/Kolkata")
    time = datetime.now(india).strftime("%Y-%m-%d %H:%M:%S")
    ws.append([name, action, time])
    wb.save(FILE)


# 🌐 HOME PAGE
@app.route("/")
def home():
    return """
    <html>
    <head>
        <title>Secure Attendance App</title>
        <style>
            body{
                font-family:Arial;
                background:#f5f7fa;
                display:flex;
                justify-content:center;
                align-items:center;
                height:100vh;
                margin:0;
            }
            .card{
                background:white;
                padding:25px;
                border-radius:15px;
                width:320px;
                text-align:center;
                box-shadow:0 4px 15px rgba(0,0,0,0.15);
            }
            select,input,button{
                width:100%;
                padding:12px;
                margin-top:10px;
                border-radius:10px;
                border:none;
                font-size:16px;
            }
            .in{background:#28a745;color:white;}
            .out{background:#dc3545;color:white;}
        </style>
    </head>

    <body>
        <div class="card">
            <h2>Secure Attendance</h2>

            <form action="/mark" method="post">
                <select name="name">
                    <option>Gaurav</option>
                    <option>Arshi</option>
                </select>

                <input type="password" name="pin" placeholder="Enter PIN" required>

                <button class="in" name="action" value="Check In">Check In</button>
                <button class="out" name="action" value="Check Out">Check Out</button>
            </form>
        </div>
    </body>
    </html>
    """


# 🔐 ATTENDANCE MARK API
@app.route("/mark", methods=["POST"])
def mark():
    name = request.form["name"]
    action = request.form["action"]
    pin = request.form["pin"]

    # SECURITY CHECK
    if name not in USERS or USERS[name] != pin:
        return """
        <script>
            alert('❌ Wrong PIN! Access Denied');
            window.location.href = "/";
        </script>
        """

    # SAVE ATTENDANCE
    save_data(name, action)

    return """
    <script>
        alert('✔ Attendance Saved Successfully');
        window.location.href = "/";
    </script>
    """


# 🚀 RUN SERVER
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
