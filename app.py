from flask import Flask, render_template, request, redirect, session
from waitress import serve
from werkzeug.security import generate_password_hash, check_password_hash
import pandas as pd
import random
import time
import json
import os
import subprocess
import smtplib
from email.mime.text import MIMEText
from datetime import datetime, timedelta



app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "supersecretkey123")

# ===============================
# JSON HELPERS (ADD HERE)
# ===============================
def load_json(file):
    if not os.path.exists(file):
        return []
    with open(file, "r") as f:
        return json.load(f)

def save_json(file, data):
    with open(file, "w") as f:
        json.dump(data, f, indent=4)
        
def add_notification(message, user=None):
    notes = load_json("notifications.json")

    notes.append({
        "message": message,
        "user": user,
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    })

    save_json("notifications.json", notes)
    
    
def create_user(username, password):
    users = load_json("users.json")

    users.append({
        "username": username,
        "password": password,

        # ✅ AUTO DEFAULT PERMISSIONS
        "login": "yes",
        "resources": "yes",
        "qa_hub": "yes",
        "help": "yes"
    })

    save_json("users.json", users)
        
        
EMAIL_ADDRESS = "your_email@gmail.com"
EMAIL_PASSWORD = "your_app_password"

EXCEL_FILE = "users.xlsx"



def send_email(to_email, subject, body):
    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = EMAIL_ADDRESS
    msg["To"] = to_email

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        server.send_message(msg)
 #-------------------------------------------------     
SUPER_ADMIN_USERNAME = os.getenv("SUPER_ADMIN_USERNAME", "dinesh_admin")

def is_super_admin():
    return session.get("user") == SUPER_ADMIN_USERNAME


# -----------------------------
# ADD HERE (STEP 1)
# -----------------------------
def is_admin():
    return session.get("role") == "admin" or is_super_admin()
    
# ✅ ADD HERE
def normalize_user_columns():
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name="users", dtype=str)
    except:
        return

    required_cols = [
        "username","password","role","purpose","approved",
        "resources_access","help_access","login_access",
        "email","must_change_password","otp",
        "otp_time","login_attempts","lock_until"
    ]

    # Ensure all required columns exist
    for col in required_cols:
        if col not in df.columns:
            df[col] = ""

    # Reorder columns
    df = df[required_cols]

    # Save back to Excel
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name="users", index=False)


#----------------------------------------------

# ✅ ADD HERE
def normalize_admin_columns():
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name="admins", dtype=str)
    except:
        return

    # ✅ Clean admin structure (removed unnecessary columns)
    required_cols = [
        "username","password","role","approved",
        "email","must_change_password","otp",
        "otp_time","login_attempts","lock_until"
    ]

    # Ensure all required columns exist
    for col in required_cols:
        if col not in df.columns:
            df[col] = ""

    # Reorder columns
    df = df[required_cols]

    # Save back to Excel
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name="admins", index=False)
        
def normalize_notification_sheet():
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name="notifications", dtype=str)
    except:
        df = pd.DataFrame(columns=["username","message"])

    if "username" not in df.columns:
        df["username"] = ""

    if "message" not in df.columns:
        df["message"] = ""

    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name="notifications", index=False)


# ✅ KEEP THIS OUTSIDE FUNCTION (no indent)
normalize_user_columns()
normalize_admin_columns()
normalize_notification_sheet()


def log_action(username, action):

    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name="logs", dtype=str)
    except:
        df = pd.DataFrame(columns=["username","action","time"])

    from datetime import datetime

    new = pd.DataFrame({
        "username":[username],
        "action":[action],
        "time":[datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
    })

    df = pd.concat([df,new], ignore_index=True)

    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name="logs", index=False)
        
#-------------------------------------------


#--------------------------
#upload
#--------------------------
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# =====================================================
# ✅ QA HUB + APPROVAL + FULL CONTROL (FINAL STABLE)
# =====================================================

QA_HUB_FILE = "qa_hub.json"


# ---------------- LOAD / SAVE ----------------
def load_folders():
    try:
        return pd.read_json(QA_HUB_FILE).to_dict(orient="records")
    except:
        return []
        
def load_excel_to_lob(folder_name):
    import pandas as pd
    import os

    file_path = os.path.join(os.getcwd(), "renewal.xlsx")

    print("📂 Excel Path:", file_path)

    if not os.path.exists(file_path):
        print("❌ Excel file NOT FOUND")
        return {}

    xls = pd.ExcelFile(file_path)

    print("📄 Sheets Found:", xls.sheet_names)

    lob_data = {}

    for sheet in xls.sheet_names:
        df = xls.parse(sheet)
        df = df.fillna("")

        records = df.to_dict(orient="records")

        lob_data[sheet.strip()] = records

    return lob_data


def save_folders(data):
    df = pd.DataFrame(data)
    df.to_json(QA_HUB_FILE, orient="records", indent=4)


# ---------------- FIND FOLDER ----------------
def get_folder(folder):
    folders = load_folders()
    return next((f for f in folders if f["name"] == folder), None)


# ---------------- MAIN PAGE ----------------
@app.route("/qa_hub", methods=["GET","POST"])
def qa_hub():

    # 🔐 LOGIN CHECK
    if "user" not in session:
        return redirect("/login/user")

    username = session["user"]

    # 🔒 ADMIN / SUPER ADMIN FULL ACCESS (FIXED)
    if not is_admin():

        users_df = pd.read_excel(EXCEL_FILE, sheet_name="users", dtype=str)

        user_data = users_df[users_df["username"] == username]

        # ✅ SAFE CHECK
        if user_data.empty:
            return redirect("/dashboard")

        user_row = user_data.iloc[0]

        # ✅ PERMISSION CHECK (SAFE FALLBACK)
        qa_access = str(
            user_row.get("qa_hub", user_row.get("resources_access", "yes"))
        ).lower()

        if qa_access != "yes":
            return "Access Denied"

    # 🔹 LOAD FOLDERS
    folders = load_folders()

    # ---------- CREATE ----------
    if request.method == "POST":

        if not is_admin():
            return "Only admin can create folders"

        folder_name = request.form["folder"].strip()

        if not folder_name:
            return "Invalid folder name"

        # prevent duplicate (case insensitive)
        if any(f["name"].lower() == folder_name.lower() for f in folders):
            return "Folder already exists"

        folders.append({
            "name": folder_name,
            "approved": "no",
            "status": "active",
            "created_by": session.get("user"),
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "data": []
        })

        save_folders(folders)

    # ---------- USER FILTER ----------
    if not is_admin():
        folders = [
            f for f in folders
            if f.get("approved") == "yes"
            and f.get("status","active") == "active"
        ]

    return render_template("qa_hub.html", folders=folders)

# =====================================================
# ✅ APPROVE FOLDER
# =====================================================
@app.route("/approve_folder/<folder>")
def approve_folder(folder):

    if not is_admin():
        return redirect("/login/admin")

    folders = load_folders()

    for f in folders:
        if f["name"] == folder:
            f["approved"] = "yes"

    save_folders(folders)

    return redirect("/qa_hub")


# =====================================================
# ✅ HIDE / SHOW FOLDER (USER ONLY IMPACT)
# =====================================================
@app.route("/toggle_folder/<folder>")
def toggle_folder(folder):

    if not is_admin():
        return redirect("/login/admin")

    folders = load_folders()

    for f in folders:
        if f["name"] == folder:
            f["status"] = "hidden" if f.get("status","active") == "active" else "active"

    save_folders(folders)

    return redirect("/qa_hub")


# =====================================================
# ✅ DELETE FOLDER (SAFE)
# =====================================================
@app.route("/delete_folder/<folder>")
def delete_folder(folder):

    if not is_admin():
        return redirect("/login/admin")

    folders = load_folders()

    folders = [f for f in folders if f["name"] != folder]

    save_folders(folders)

    return redirect("/qa_hub")


# =====================================================
# ✅ QA FOLDER PAGE
# =====================================================
from datetime import datetime, timedelta

@app.route("/qa_hub/<folder>", methods=["GET","POST"])
def qa_folder(folder):

    if "user" not in session:
        return redirect("/login/user")

    folders = load_folders()
    folder_obj = get_folder(folder)

    if not folder_obj:
        return "Folder not found"

    # ✅ PERMISSION CONTROL
    can_view = True
    if not is_admin():
        if folder_obj.get("approved") != "yes":
            can_view = False
        if folder_obj.get("status", "active") != "active":
            can_view = False

    # ✅ GET DATA (RENEWAL)
    data = folder_obj.get("lob_data", {})

    # 🔥 DELINQUENCY DATA LOAD (SAFE)
    import os, json

    delinquency_data = {}

    if "delinquency" in folder.lower():
        try:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            path = os.path.join(base_dir, "data", "delinquency.json")

            print("📂 Loading Delinquency JSON from:", path)

            if os.path.exists(path):
                with open(path) as f:
                    delinquency_data = json.load(f)
                print("✅ Loaded successfully")
            else:
                print("❌ File NOT found at:", path)

        except Exception as e:
            print("❌ Error loading JSON:", str(e))
            delinquency_data = {}


    # ✅ LOAD FROM EXCEL ONLY FOR NON-DELINQUENCY
        if "delinquency" not in folder.lower():
            if not data or len(data) == 0:
                data = load_excel_to_lob(folder)

        if data:
            for f in folders:
                if f["name"] == folder:
                    f["lob_data"] = data
            save_folders(folders)

    # ✅ HANDLE ADMIN ACTIONS
    if request.method == "POST":

        if not is_admin():
            return "Only admin allowed"

        action = request.form.get("action")

        if action == "add_lob":
            lob = request.form.get("lob")
            if lob and lob not in data:
                data[lob] = []

        elif action == "add_row":
            lob = request.form.get("lob")
            row = {}

            for k in request.form:
                if k not in ["action","lob"]:
                    row[k] = request.form.get(k)

            if lob in data:
                data[lob].append(row)

        for f in folders:
            if f["name"] == folder:
                f["lob_data"] = data

        save_folders(folders)

        return redirect(f"/qa_hub/{folder}")

    # ✅ FINAL RETURN (ONLY ONCE)
    # 🔥 DELINQUENCY SEPARATE PAGE
        # ✅ FINAL RETURN (ONLY ONCE)

    # 🔥 DELINQUENCY PAGE
    if "delinquency" in folder.lower():
        return render_template(
            "delinquency.html",
            folder=folder,
            delinquency_data=delinquency_data,
            is_admin=is_admin(),
            can_view=can_view
        )

    # 🔥 DEFAULT (RENEWAL PAGE)
    return render_template(
        "qa_folder.html",
        folder=folder,
        data=data,
        is_admin=is_admin(),
        can_view=can_view
    )


# =====================================================
# ✅ DELETE ENTRY (FUTURE SAFE)
# =====================================================
@app.route("/delete_entry/<folder>/<entry_id>")
def delete_entry(folder, entry_id):

    if not is_admin():
        return redirect("/login/admin")

    folders = load_folders()

    for f in folders:
        if f["name"] == folder:
            f["data"] = [d for d in f.get("data", []) if d.get("id") != entry_id]

    save_folders(folders)

    return redirect(f"/qa_hub/{folder}")


# =====================================================
# ✅ RAISE QUERY (NO CHANGE NEEDED)
# =====================================================
@app.route("/raise_query", methods=["GET","POST"])
def raise_query():

    if "user" not in session:
        return redirect("/login/user")

    folders = load_folders()

    folder_list = [
        f["name"]
        for f in folders
        if f.get("approved") == "yes" and f.get("status") == "active"
    ]

    selected_folder = request.args.get("folder", "")

    if request.method == "POST":

        query = {
            "user": session["user"],
            "folder": request.form.get("folder"),
            "topic": request.form.get("topic"),
            "type": request.form.get("type"),
            "message": request.form.get("description"),
            "status": "pending"
        }

        data = load_json("help_requests.json")
        data.append(query)
        save_json("help_requests.json", data)

        # ✅ NOTIFY ADMIN
        add_notification(f"📩 New query from {session['user']} in {query['folder']}")

        return redirect("/qa_hub")

    return render_template(
        "raise_query.html",
        folders=folder_list,
        selected_folder=selected_folder
    )
    
    
    
    
 #---------------------------------------------------------------------------------------------------------------------------

# QA Quotes for welcome page
qa_quotes = [
    "The role of QA is not to find bugs, but to prevent bugs from being created in the first place.",
    "Quality is free, but it's not a gift.",
    "If a thing's worth doing, it's worth doing well.",
    "Quality is not an act, it is a habit. – Aristotle",
    "Quality is everyone's responsibility. — W. Edwards Deming",
    "Quality means doing it right when no one is looking.",
    "The more we sweat in testing, the less we bleed in production.",
    "Be a yardstick of quality. Some people aren't used to an environment where excellence is expected.",
    "Quality is never an accident; it is always the result of intelligent effort.",
    "Good software testers break things, but great testers prevent things from being broken."
]

# -----------------------------
# Welcome Page
# -----------------------------
# -----------------------------
@app.route("/")
def welcome():
    quote = random.choice(qa_quotes)
    return render_template("welcome.html", quote=quote)

# -----------------------------
# Login Route (admin/user separate)
# -----------------------------
# -----------------------------
# Login Route (admin/user separate)
# -----------------------------
@app.route("/login/<string:login_type>", methods=["GET", "POST"])
def login_route(login_type):
    if request.method == "POST":
        username = request.form["username"].strip()
        password = request.form["password"].strip()

        # ---------------- ADMIN LOGIN ----------------
        if login_type == "admin":

            try:
                admins = pd.read_excel(EXCEL_FILE, sheet_name="admins", dtype=str)
                admins["username"] = admins["username"].str.strip()
                admins["password"] = admins["password"].str.strip()
            except:
                admins = pd.DataFrame(columns=[
                    "username","password","role","approved",
                    "email","must_change_password","otp",
                    "otp_time","login_attempts","lock_until"
                ])

            admin_user = admins[admins.username == username]

            # 🔥 SUPER ADMIN (WITH SECURITY)
            if username == SUPER_ADMIN_USERNAME:
                if not admin_user.empty:

                    # 🔐 LOGIN ATTEMPT CHECK
                    raw_attempts = admin_user.iloc[0].get("login_attempts", "0")
                    try:
                        attempts = int(float(raw_attempts)) if raw_attempts not in [None,"","nan"] else 0
                    except:
                        attempts = 0

                    lock_until = admin_user.iloc[0].get("lock_until","")

                    if lock_until:
                        try:
                            lock_time = datetime.strptime(lock_until,"%Y-%m-%d %H:%M:%S")
                            if datetime.now() < lock_time:
                                return "Admin account locked. Try later."
                        except:
                            pass

                    stored_password = admin_user.iloc[0]["password"]

                    if not check_password_hash(stored_password, password):

                        attempts += 1
                        df = pd.read_excel(EXCEL_FILE, sheet_name="admins", dtype=str)
                        df.loc[df["username"] == username, "login_attempts"] = str(attempts)

                        if attempts >= 3:
                            lock_time = datetime.now() + timedelta(minutes=10)
                            df.loc[df["username"] == username, "lock_until"] = lock_time.strftime("%Y-%m-%d %H:%M:%S")

                        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                            df.to_excel(writer, sheet_name="admins", index=False)

                        return "Invalid Admin Password"

                    # ✅ RESET AFTER SUCCESS
                    df = pd.read_excel(EXCEL_FILE, sheet_name="admins", dtype=str)
                    df.loc[df["username"] == username, "login_attempts"] = "0"
                    df.loc[df["username"] == username, "lock_until"] = ""

                    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        df.to_excel(writer, sheet_name="admins", index=False)

                    session["user"] = username
                    session["role"] = "admin"
                    log_action(username, "super admin login")
                    
                # 👇 ADD THIS
                    add_notification(f"👑 You logged in as Super Admin", username)

                    return redirect("/dashboard")

                else:
                    return "Super Admin not found in admin sheet"

            # ---------------- NORMAL ADMIN ----------------
            if not admin_user.empty:

                # 🔐 LOGIN ATTEMPT CHECK
                raw_attempts = admin_user.iloc[0].get("login_attempts", "0")
                try:
                    attempts = int(float(raw_attempts)) if raw_attempts not in [None,"","nan"] else 0
                except:
                    attempts = 0

                lock_until = admin_user.iloc[0].get("lock_until","")

                if lock_until:
                    try:
                        lock_time = datetime.strptime(lock_until,"%Y-%m-%d %H:%M:%S")
                        if datetime.now() < lock_time:
                            return "Admin account locked. Try later."
                    except:
                        pass

                stored_password = admin_user.iloc[0]["password"]

                # ❌ WRONG PASSWORD
                if not check_password_hash(stored_password, password):

                    attempts += 1

                    df = pd.read_excel(EXCEL_FILE, sheet_name="admins", dtype=str)
                    df.loc[df["username"] == username, "login_attempts"] = str(attempts)

                    if attempts >= 3:
                        lock_time = datetime.now() + timedelta(minutes=10)
                        df.loc[df["username"] == username, "lock_until"] = lock_time.strftime("%Y-%m-%d %H:%M:%S")

                    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        df.to_excel(writer, sheet_name="admins", index=False)

                    return "Invalid Admin Password"

                # ✅ RESET AFTER SUCCESS
                df = pd.read_excel(EXCEL_FILE, sheet_name="admins", dtype=str)
                df.loc[df["username"] == username, "login_attempts"] = "0"
                df.loc[df["username"] == username, "lock_until"] = ""

                with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name="admins", index=False)

                session["user"] = username
                session["role"] = "admin"
                if is_super_admin():
                    log_action(username, "super admin login")
                else:
                    log_action(username, "admin login")
                 
                 # 👇 ADD THIS
                add_notification(f"🛠 You logged in as Admin", username)


                approved = str(admin_user.iloc[0].get("approved","no")).strip().lower()
                if approved == "yes" or is_super_admin():
                    return redirect("/dashboard")
                else:
                    return redirect("/admin_approval")

            return "Invalid Admin Login"

        # ---------------- USER LOGIN ----------------
        else:
            try:
                users = pd.read_excel(EXCEL_FILE, sheet_name="users", dtype=str)
                users["username"] = users["username"].str.strip()
                users["password"] = users["password"].str.strip()
            except:
                users = pd.DataFrame(columns=["username", "password", "role", "purpose", "approved"])

            user = users[users.username == username]

            if not user.empty:

                raw_attempts = user.iloc[0].get("login_attempts", "0")

                try:
                    attempts = int(float(raw_attempts)) if raw_attempts not in [None, "", "nan"] else 0
                except:
                    attempts = 0

                lock_until = user.iloc[0].get("lock_until", "")

                # LOCK CHECK
                if lock_until:
                    try:
                        lock_time = datetime.strptime(lock_until, "%Y-%m-%d %H:%M:%S")
                        if datetime.now() < lock_time:
                            return "Account locked. Try after some time."
                    except:
                        pass

                stored_password = user.iloc[0]["password"]

                # WRONG PASSWORD
                if not check_password_hash(stored_password, password):

                    attempts += 1

                    df = pd.read_excel(EXCEL_FILE, sheet_name="users", dtype=str)
                    df.loc[df["username"] == username, "login_attempts"] = str(attempts)

                    if attempts >= 3:
                        lock_time = datetime.now() + timedelta(minutes=10)
                        df.loc[df["username"] == username, "lock_until"] = lock_time.strftime("%Y-%m-%d %H:%M:%S")

                    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        df.to_excel(writer, sheet_name="users", index=False)

                    return "Invalid password"

                # FORCE PASSWORD CHANGE
                must_change = str(user.iloc[0].get("must_change_password", "no")).lower()

                if must_change == "yes":
                    session["reset_user"] = username
                    return redirect("/change_password")

                # RESET ATTEMPTS
                df = pd.read_excel(EXCEL_FILE, sheet_name="users", dtype=str)
                df.loc[df["username"] == username, "login_attempts"] = "0"
                df.loc[df["username"] == username, "lock_until"] = ""

                with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name="users", index=False)

                # NORMAL LOGIN
                session["user"] = username
                session["role"] = user.iloc[0]["role"]
                log_action(username, "login success")
                
                # 👇 ADD THIS
                add_notification(f"👤 You logged in", username)

                approved = str(user.iloc[0].get("approved","no")).strip().lower()

                login_access = str(user.iloc[0].get("login_access", "yes")).strip().lower()
                if not is_super_admin():
                    if login_access != "yes":
                        return "Your login access has been disabled by admin."

                if approved == "yes" or is_super_admin():
                    return redirect("/dashboard")
                else:
                    return "Login failed. Your account is not approved yet."

            return "Invalid User Login"

    return render_template("login.html", login_type=login_type)


# -----------------------------
# Forgot Password
# -----------------------------
@app.route("/forgot_password", methods=["GET","POST"])
def forgot_password():

    if request.method == "POST":

        username = request.form["username"].strip()
        email = request.form["email"].strip()

        try:
            df = pd.read_excel(EXCEL_FILE, sheet_name="users", dtype=str)
        except:
            return "User data not found"

        user = df[df["username"] == username]

        if user.empty:
            return "User not found"

        stored_email = str(user.iloc[0].get("email","")).strip()

        if email != stored_email:
            return "Email does not match"

        from datetime import datetime

        # 🔢 Generate OTP
        otp = str(random.randint(100000, 999999))

        df.loc[df["username"] == username, "otp"] = otp
        df.loc[df["username"] == username, "otp_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # ✅ SEND EMAIL
        try: 
            send_email(email, "Password Reset OTP", f"Your OTP is: {otp}")
        except:
            print("Email failed, continuing...")
        

        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name="users", index=False)

        return render_template("verify_otp.html", username=username)

    return render_template("forgot_password.html")
#--------------------------------------------------------

@app.route("/change_password", methods=["GET","POST"])
def change_password():

    if "reset_user" not in session:
        return redirect("/login/user")

    if request.method == "POST":

        new_password = request.form["new_password"].strip()

        if len(new_password) < 6:
            return "Password must be at least 6 characters"

        username = session["reset_user"]

        df = pd.read_excel(EXCEL_FILE, sheet_name="users", dtype=str)

        df.loc[df["username"] == username, "password"] = generate_password_hash(new_password)
        df.loc[df["username"] == username, "must_change_password"] = "no"

        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name="users", index=False)

        session.pop("reset_user", None)

        return redirect("/login/user")

    return render_template("change_password.html")
    
    
  #---------------------------------------------------
@app.route("/verify_otp", methods=["POST"])
def verify_otp():

    username = request.form["username"]
    entered_otp = request.form["otp"]

    df = pd.read_excel(EXCEL_FILE, sheet_name="users", dtype=str)

    user = df[df["username"] == username]

    if user.empty:
        return "User not found"

    # ✅ GET STORED VALUES
    stored_otp = str(user.iloc[0].get("otp", "")).strip()
    otp_time = user.iloc[0].get("otp_time", "")

    # ❌ INVALID OTP
    if entered_otp != stored_otp:
        return "Invalid OTP"

    # ✅ EXPIRY CHECK (SAFE)
    if otp_time:
        try:
            otp_time = datetime.strptime(otp_time, "%Y-%m-%d %H:%M:%S")
            if datetime.now() > otp_time + timedelta(minutes=5):
                return "OTP expired"
        except:
            return "OTP error. Try again."

    # ✅ LOG ACTION
    log_action(username, "password reset via OTP")

    # 🔑 Generate temporary password
    temp_password = "Temp@" + str(random.randint(1000,9999))

    df.loc[df["username"] == username, "password"] = generate_password_hash(temp_password)
    df.loc[df["username"] == username, "must_change_password"] = "yes"
    df.loc[df["username"] == username, "otp"] = ""
    df.loc[df["username"] == username, "otp_time"] = ""

    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name="users", index=False)

    return f"✅ Your temporary password is: {temp_password}"
    
    
    
 #------------------------------
    
@app.route("/admin_reset_password/<username>")
def admin_reset_password(username):
    
    if username == SUPER_ADMIN_USERNAME and not is_super_admin():
        return "Not allowed"

    if "user" not in session or not is_admin():
        return redirect("/login/admin")

    df = pd.read_excel(EXCEL_FILE, sheet_name="users", dtype=str)

    # ✅ FIX: check user exists
    if username not in df["username"].values:
        return "User not found"

    temp_password = "Admin@" + str(random.randint(1000,9999))

    df.loc[df["username"] == username, "password"] = generate_password_hash(temp_password)
    df.loc[df["username"] == username, "must_change_password"] = "yes"
    df.loc[df["username"] == username, "login_attempts"] = "0"
    df.loc[df["username"] == username, "lock_until"] = ""

    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name="users", index=False)

    log_action(session.get("user"), f"reset password for {username}")

    return f"Temp password for {username}: {temp_password}"

# -----------------------------
# Register Route
# -----------------------------
@app.route("/register", methods=["GET","POST"])
def register_route():

    message = ""

    if request.method == "POST":

        username = request.form["username"].strip()
        password = request.form["password"].strip()
        role = request.form["role"].strip()
        purpose = request.form.get("purpose","")
        # ✅ ADD THIS
        email = request.form["email"].strip()

        # 🔒 Password validation
        if len(password) < 6:
            message = "Password must be at least 6 characters"
            return render_template("register.html", message=message)

        sheet_name = "admins" if role == "admin" else "users"

        try:
            df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, dtype=str)
        except:
            if sheet_name == "admins":
                df = pd.DataFrame(columns=["username","password","role","approved","email","must_change_password","otp",
                                "otp_time","login_attempts","lock_until"])
            else:
                df = pd.DataFrame(columns=[
                        "username","password","role","purpose","approved",
                         "login_access","resources_access","help_access",
                            "email","must_change_password","otp"
])

        # ❌ Duplicate check
        if username in df["username"].values:
            message = "Username already exists"
            return render_template("register.html", message=message)

        # ✅ Default permissions
        new_entry = {
            "username": username,
            "password": generate_password_hash(password),
            "role": role,
            "approved": "no",
            "must_change_password": "no",
            "otp": "",
            "email": email,   # ✅ ADD THIS
            "otp_time": "",
            "login_attempts": "0",   # ✅ ADD HERE
            "lock_until": ""         # ✅ ADD HERE
        }

        if sheet_name == "users":
            new_entry["purpose"] = purpose
            new_entry["login_access"] = "yes"
            new_entry["resources_access"] = "yes"
            new_entry["help_access"] = "yes"
            new_entry["qa_hub"] = "yes"

        df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)

        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

        message = "Account created! Waiting for admin approval."

    return render_template("register.html", message=message)

# -----------------------------
# Dashboard
# -----------------------------
@app.route("/dashboard")
def dashboard_route():

    if "user" not in session:
        return redirect("/login/user")

    # Default values
    total_resources = 0
    help_requests = 0
    users_count = 0
    recent_resources = []
    recent_help = []
    top_folder = "N/A"

    try:
        resources_df = pd.read_excel(EXCEL_FILE, sheet_name="resources", dtype=str)
        total_resources = len(resources_df)
        recent_resources = resources_df["resource_name"].dropna().tail(3).tolist()
    except:
        pass

    try:
        help_df = pd.read_excel(EXCEL_FILE, sheet_name="help_requests", dtype=str)
        help_requests = len(load_json("help_requests.json"))
        recent_help = help_df["message"].dropna().tail(3).tolist()
    except:
        pass

    try:
        users_df = pd.read_excel(EXCEL_FILE, sheet_name="users", dtype=str)
        users_count = len(users_df)
    except:
        pass

    try:
        resources_df["folder"] = resources_df["file"].str.split("/").str[0]
        top_folder = resources_df["folder"].mode()[0]
    except:
        pass

    # ✅ PERMISSIONS FIX
    try:
        users_df = pd.read_excel(EXCEL_FILE, sheet_name="users", dtype=str)

        user_data = users_df[users_df["username"] == session["user"]]

        if not user_data.empty:
            user_row = user_data.iloc[0]

            user_permissions = {
                "qa_hub": str(user_row.get("qa_hub", user_row.get("resources_access", "yes"))).lower(),
                "resources": str(user_row.get("resources_access", "yes")).lower(),
                "help": str(user_row.get("help_access", "yes")).lower(),
                "login": str(user_row.get("login_access", "yes")).lower()
            }
        else:
            user_permissions = {
                "qa_hub": "yes",
                "resources": "yes",
                "help": "yes",
                "login": "yes"
            }

    except:
        user_permissions = {
            "qa_hub": "yes",
            "resources": "yes",
            "help": "yes",
            "login": "yes"
        }

    # ✅ MUST BE OUTSIDE TRY/EXCEPT
    return render_template(
        "dashboard.html",
        user=session["user"],
        role=session["role"],
        total_resources=total_resources,
        help_requests=help_requests,
        users_count=users_count,
        recent_resources=recent_resources,
        recent_help=recent_help,
        top_folder=top_folder,
        user_permissions=user_permissions
    )

#------------------
# @app.route("/dashboard")
# def dashboard_route():
#    if "user" not in session:
#       return redirect("/login/user")
#    return render_template("dashboard.html", user=session["user"], role=session["role"])
#-------------------

# -----------------------------
# Logout
# -----------------------------
@app.route("/logout")
def logout_route():
    session.pop("user", None)
    session.pop("role", None)
    return redirect("/")

# -----------------------------
# Admin Approval Panel
# -----------------------------
@app.route("/admin_approval", methods=["GET","POST"])

def admin_approval_route():

    if "user" not in session or not is_admin():
        return redirect("/login/admin")

    if request.method == "POST":
        new_username = request.form["username"].strip()
        new_password = request.form["password"].strip()
        new_role = request.form["role"].strip()
        new_purpose = request.form.get("purpose","")
        new_email = request.form["email"].strip()

        sheet_name = "admins" if new_role == "admin" else "users"

        try:
            df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, dtype=str)

        except:
            if sheet_name == "admins":
                df = pd.DataFrame(columns=["username","password","role","approved","email","must_change_password","otp",
                                "otp_time","login_attempts","lock_until"])

            else:
                # ✅ FIXED (removed wrong nested df)
                df = pd.DataFrame(columns=[
                    "username","password","role","purpose","approved",
                    "login_access","resources_access","help_access",
                    "email","must_change_password","otp",
                    "otp_time","login_attempts","lock_until"
                ])

        # ✅ NEW ENTRY
        new_entry = {
            "username": new_username,
            "password": generate_password_hash(new_password),
            "role": new_role,
            "approved": "no",
            "email": new_email,
            "must_change_password": "no",
            "otp": "",
            "otp_time": "",
            "login_attempts": "0",
            "lock_until": ""
        }

        if sheet_name == "users":
            new_entry["purpose"] = new_purpose
            new_entry["login_access"] = "yes"
            new_entry["resources_access"] = "yes"
            new_entry["help_access"] = "yes"
            
        # 🚨 ADD THIS LINE HERE
        if new_username in df["username"].values:
            return "Username already exists"

        df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)

        try:
            with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        except PermissionError:
            return "Close users.xlsx and try again."

    # ---------------- LOAD DATA ----------------
    try:
        users = pd.read_excel(EXCEL_FILE, sheet_name="users", dtype=str)
    except:
        users = pd.DataFrame(columns=[
            "username","password","role","purpose","approved"
        ])

    try:
        admins = pd.read_excel(EXCEL_FILE, sheet_name="admins", dtype=str)
    except:
        admins = pd.DataFrame(columns=[
            "username","password","role","approved"
        ])

    return render_template(
        "admin_approval.html",
        users=users.to_dict(orient="records"),
        admins=admins.to_dict(orient="records"),
        super_admin=SUPER_ADMIN_USERNAME # ✅ comma added
    )
                           
                           
# -----------------------------
# Admin User Permission Panel
# -----------------------------
@app.route("/admin_user_permissions")
def admin_user_permissions():

    if "user" not in session or not is_admin():
        return redirect("/login/admin")

    try:
        users_df = pd.read_excel(EXCEL_FILE, sheet_name="users", dtype=str)
    except:
        users_df = pd.DataFrame()

    users = users_df.to_dict(orient="records")

    return render_template("admin_user_permissions.html", users=users)
    
    
    
@app.route("/toggle_user_permission/<username>/<permission>")
def toggle_user_permission(username, permission):

    # 🔐 SESSION CHECK
    if "user" not in session or not is_admin():
        return redirect("/login/admin")

    # 🔒 BLOCK NORMAL ADMIN FROM TOUCHING SUPER ADMIN
    if username == SUPER_ADMIN_USERNAME and not is_super_admin():
        return "Not allowed"

    # 🔒 EVEN SUPER ADMIN SHOULD NOT BREAK HIMSELF
    if username == SUPER_ADMIN_USERNAME:
        return redirect("/dashboard")

    users_df = pd.read_excel(EXCEL_FILE, sheet_name="users", dtype=str)

    # ✅ HANDLE USER NOT FOUND (IMPORTANT FIX)
    if username not in users_df["username"].values:
        return "User not found"

    # ✅ AUTO CREATE COLUMN IF NOT EXISTS
    if permission not in users_df.columns:
        users_df[permission] = "yes"

    # ✅ HANDLE EMPTY / NaN VALUE
    current_value = users_df.loc[
        users_df["username"] == username, permission
    ].values[0]

    current = str(current_value).strip().lower() if pd.notna(current_value) else "yes"

    # ✅ TOGGLE LOGIC
    new_value = "no" if current == "yes" else "yes"

    users_df.loc[
        users_df["username"] == username, permission
    ] = new_value

    # ✅ SAVE BACK
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        users_df.to_excel(writer, sheet_name="users", index=False)

    return redirect("/admin_user_permissions")
# -----------------------------
# Admin Help Requests
# -----------------------------
# -----------------------------
# Admin Help Requests (FIXED)
# -----------------------------
@app.route("/admin_help")
def admin_help():

    if "user" not in session or not is_admin():
        return redirect("/login/admin")

    requests = load_json("help_requests.json")

    return render_template("admin_help.html", requests=requests)


# -----------------------------
# Approve / Reject Request (FIXED)
# -----------------------------
@app.route("/update_request/<int:index>/<action>")
def update_request(index, action):

    if "user" not in session or not is_admin():
        return redirect("/login/admin")

    data = load_json("help_requests.json")

    if index < len(data):

        req = data[index]
        req["status"] = action

        save_json("help_requests.json", data)

        username = req.get("user", "")
        topic = req.get("topic", "Request")

        # ✅ USER NOTIFICATION
        if action == "approve":
            add_notification(
                f"✅ Your request '{topic}' has been approved. Thanks for raising it!",
                username
            )
        else:
            add_notification(
                f"❌ Your request '{topic}' was rejected. Please review and resubmit if needed.",
                username
            )

        # ✅ GLOBAL NOTIFICATION (ALL USERS)
        add_notification(
            f"📢 Admin updated request '{topic}'. Please check QA Hub."
        )

    return redirect("/admin_help")
# -----------------------------
# Approve / Reject Users
# -----------------------------
@app.route("/approve_user/<string:user_type>/<string:username>/<string:action>")
def approve_user_route(user_type, username, action):
    
    if username == SUPER_ADMIN_USERNAME and not is_super_admin():
        return "Not allowed"
    
    if "user" not in session or not is_admin():
        return redirect("/login/admin")

    sheet_name = "admins" if user_type=="admin" else "users"

    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, dtype=str)
    except:
        df = pd.DataFrame()

    df.loc[df.username==username,"approved"] = "yes" if action=="approve" else "no"

    try:
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    except PermissionError:
        return "Cannot update Excel. Close it and try again."

    return redirect("/admin_approval")

# -----------------------------
# User Help Request
# -----------------------------
@app.route("/help", methods=["GET","POST"])
def help_request_route():
    if "user" not in session:
        return redirect("/login/user")

    message_sent = ""

    if request.method == "POST":

        username = session.get("user")
        users_df = pd.read_excel(EXCEL_FILE, sheet_name="users", dtype=str)

        user_row = users_df[users_df["username"] == username]

        if not is_super_admin():
            if not user_row.empty:
                if user_row.iloc[0].get("help_access","yes").lower() != "yes":
                    return "Admin has disabled help request access."

        message = request.form["message"].strip()

        try:
            help_df = pd.read_excel(EXCEL_FILE, sheet_name="help_requests", dtype=str)
        except:
            help_df = pd.DataFrame(columns=["username","message"])

        new = pd.DataFrame({"username":[username], "message":[message]})
        help_df = pd.concat([help_df,new], ignore_index=True)

        try:
            with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                help_df.to_excel(writer, sheet_name="help_requests", index=False)
        except PermissionError:
            return "Cannot save help request. Close users.xlsx if open."

        message_sent = "Your request has been submitted successfully!"

    return render_template("help.html", message_sent=message_sent)

# -----------------------------
# Admin Resources
# -----------------------------
@app.route("/admin_resources", methods=["GET","POST"])
def admin_resources_route():

    if "user" not in session or not is_admin():
        return redirect("/login/admin")

    message = ""   # ✅ popup message

    if request.method == "POST":

        resource_name = request.form["resource_name"].strip()
        link = request.form["link"].strip()
        allowed_users = request.form["allowed_users"].strip()
        blocked_users = request.form["blocked_users"].strip()

        folder = request.form.get("folder","general")

        file = request.files.get("file")
        filename = ""

        # ✅ LOAD EXISTING DATA FIRST
        try:
            resources = pd.read_excel(EXCEL_FILE, sheet_name="resources", dtype=str)
        except:
            resources = pd.DataFrame(columns=["id","resource_name","link","file","allowed_users","blocked_users","status"])

        # ✅ SAVE FILE
        if file and file.filename != "":
            os.makedirs(os.path.join(UPLOAD_FOLDER, folder), exist_ok=True)

            filename = file.filename
            file_path = os.path.join(UPLOAD_FOLDER, folder, filename)
            file.save(file_path)

            filename = f"{folder}/{filename}"

        # ✅ DUPLICATE CHECK
        if not resources.empty:
            duplicate = resources[
                (resources["resource_name"].fillna("").str.strip().str.lower() == resource_name.lower()) &
                (resources["link"].fillna("").str.strip() == link) &
                (resources["file"].fillna("").str.strip() == filename)
            ]

            if not duplicate.empty:
                return redirect("/admin_resources?msg=duplicate")

        # ✅ CREATE ID
        new_id = resources["id"].astype(int).max() + 1 if not resources.empty else 1

        # ✅ NEW ROW
        new_row = {
            "id": new_id,
            "resource_name": resource_name,
            "link": link,
            "file": filename,
            "allowed_users": allowed_users,
            "blocked_users": blocked_users,
            "status": "active"
        }

        resources = pd.concat([resources, pd.DataFrame([new_row])], ignore_index=True)

        # ✅ SAVE
        try:
            with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                resources.to_excel(writer, sheet_name="resources", index=False)
        except PermissionError:
            return "Close users.xlsx and try again."

        return redirect("/admin_resources?msg=success")

    # ✅ GET MESSAGE
    msg = str(request.args.get("msg") or "").strip()

    try:
        resources = pd.read_excel(EXCEL_FILE, sheet_name="resources", dtype=str)
    except:
        resources = pd.DataFrame()

    try:
        users_df = pd.read_excel(EXCEL_FILE, sheet_name="users", dtype=str)
        user_list = users_df["username"].dropna().tolist()
    except:
        user_list = []

    folders = []
    for f in os.listdir(UPLOAD_FOLDER):
        if os.path.isdir(os.path.join(UPLOAD_FOLDER, f)):
            folders.append(f)

    return render_template(
        "admin_resources.html",
        resources=resources.to_dict(orient="records"),
        users=user_list,
        folders=folders,
        msg=msg   # ✅ send to UI
    )
                           
                           
                           
#New - delete
@app.route("/delete_resource/<int:rid>")
def delete_resource(rid):

    if "user" not in session or not is_admin():
        return redirect("/login/admin")

    try:
        resources = pd.read_excel(EXCEL_FILE, sheet_name="resources", dtype=str)
    except:
        return redirect("/admin_resources")

    resources = resources[resources["id"].astype(int) != rid]

    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        resources.to_excel(writer, sheet_name="resources", index=False)

    return redirect("/admin_resources")
#----------------------------------
#----------------------------------
@app.route("/toggle_resource/<int:rid>")
def toggle_resource(rid):

    if "user" not in session or not is_admin():
        return redirect("/login/admin")

    resources = pd.read_excel(EXCEL_FILE, sheet_name="resources", dtype=str)

    current_status = resources.loc[
        resources["id"].astype(int) == rid, "status"
    ].values[0]

    if current_status == "active":
        resources.loc[
            resources["id"].astype(int) == rid, "status"
        ] = "disabled"
    else:
        resources.loc[
            resources["id"].astype(int) == rid, "status"
        ] = "active"

    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        resources.to_excel(writer, sheet_name="resources", index=False)

    return redirect("/admin_resources")
    
    
#---------------------
@app.route("/delete_multiple")
def delete_multiple():

    if "user" not in session or not is_admin():
        return redirect("/login/admin")

    ids = request.args.get("ids", "")
    id_list = ids.split(",")

    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name="resources", dtype=str)
    except:
        return redirect("/admin_resources")

    # ✅ REMOVE SELECTED IDS
    df = df[~df["id"].astype(str).isin(id_list)]

    # ✅ SAVE BACK
    try:
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name="resources", index=False)
    except PermissionError:
        return "Close users.xlsx and try again."

    # ✅ REDIRECT WITH MESSAGE
    return redirect("/admin_resources?msg=deleted")
    
#New edit

@app.route("/edit_resource/<int:rid>", methods=["GET","POST"])
def edit_resource(rid):

    if "user" not in session or not is_admin():
        return redirect("/login/admin")

    resources = pd.read_excel(EXCEL_FILE, sheet_name="resources", dtype=str)

    if request.method == "POST":

        resources.loc[resources["id"].astype(int)==rid,"resource_name"] = request.form["resource_name"]
        resources.loc[resources["id"].astype(int)==rid,"link"] = request.form["link"]
        resources.loc[resources["id"].astype(int)==rid,"allowed_users"] = request.form["allowed_users"]
        resources.loc[resources["id"].astype(int)==rid,"blocked_users"] = request.form["blocked_users"]

        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            resources.to_excel(writer, sheet_name="resources", index=False)

        return redirect("/admin_resources")

    resource = resources[resources["id"].astype(int)==rid].iloc[0]

    return render_template("edit_resource.html", r=resource)

# -----------------------------
# User Resources
# -----------------------------
@app.route("/resources")
def resources_route():

    if "user" not in session:
        return redirect("/login/user")

    username = session.get("user")

    # 🔥 SUPER ADMIN CHECK
    #super_admin_flag = (username == SUPER_ADMIN_USERNAME)

    # ---------------- USER ACCESS CHECK ----------------
    if not is_super_admin():
        users_df = pd.read_excel(EXCEL_FILE, sheet_name="users", dtype=str)
        user_row = users_df[users_df["username"] == username]

        if not user_row.empty:
            if str(user_row.iloc[0].get("resources_access", "yes")).lower() != "yes":
                return "Admin has disabled your resource access."

    # ---------------- LOAD RESOURCES ----------------
    try:
        resources = pd.read_excel(EXCEL_FILE, sheet_name="resources", dtype=str)
    except:
        resources = pd.DataFrame()

    visible_resources = []

    for _, r in resources.iterrows():

        status = str(r.get("status", "active")).lower()

        if status != "active":
            continue

        # 🔥 SUPER ADMIN SEES EVERYTHING
        if not is_super_admin():

            allowed = str(r.get("allowed_users", "")).split(",")
            blocked = str(r.get("blocked_users", "")).split(",")

            allowed = [a.strip() for a in allowed if a.strip()]
            blocked = [b.strip() for b in blocked if b.strip()]

            if username not in allowed:
                continue

            if username in blocked:
                continue

        visible_resources.append(r)

    # ---------------- GROUP BY FOLDER ----------------
    grouped = {}

    for r in visible_resources:

        folder = str(r.get("file", "general")).split("/")[0]

        if folder not in grouped:
            grouped[folder] = []

        grouped[folder].append(r)

    return render_template("resources.html", grouped=grouped)

#------------view route
@app.route("/view_resource/<int:rid>")
def view_resource(rid):

    if "user" not in session:
        return redirect("/login/user")

    username = session.get("user")

    try:
        resources = pd.read_excel(EXCEL_FILE, sheet_name="resources", dtype=str)
    except:
        return "Resource error"

    resource = resources[resources["id"].astype(int) == rid]

    if resource.empty:
        return "Resource not found"

    r = resource.iloc[0]

    allowed = str(r.get("allowed_users","")).split(",")
    blocked = str(r.get("blocked_users","")).split(",")

    allowed = [a.strip() for a in allowed if a.strip()]
    blocked = [b.strip() for b in blocked if b.strip()]

    # ✅ SUPER ADMIN BYPASS # ✅ SUPER ADMIN + ADMIN BYPASS
    if not (is_super_admin() or session.get("role") == "admin"):

        if username in blocked:
            return "You are not allowed to view this resource."

        if allowed and username not in allowed:
            return "You are not allowed to view this resource."

    return render_template(
        "secure_viewer.html",
        file=r.get("file",""),
        link=r.get("link",""),
        user=username
    )







@app.route("/convert_docx/<path:file>")
def convert_docx(file):

    docx_path = os.path.join(UPLOAD_FOLDER, file)
    pdf_path = docx_path.replace(".docx", ".pdf")

    # Check if source file exists
    if not os.path.exists(docx_path):
        return "Source file not found"

    # Convert only if PDF not already created
    if not os.path.exists(pdf_path):

        try:
            result = subprocess.run([
                "soffice",
                "--headless",
                "--convert-to",
                "pdf",
                docx_path,
                "--outdir",
                os.path.dirname(docx_path)
            ], capture_output=True)

            # ❌ Conversion failed
            if result.returncode != 0:
                print(result.stderr.decode())
                return "Conversion failed"

        except Exception as e:
            print(str(e))
            return "Cannot convert document."

    pdf_file = file.replace(".docx", ".pdf")

    return redirect(f"/uploads/{pdf_file}")




from flask import send_from_directory


#--------------------------------
@app.route('/uploads/<path:filename>')
def uploaded_file(filename):

    if "user" not in session:
        return redirect("/login/user")

    username = session.get("user")

    try:
        resources = pd.read_excel(EXCEL_FILE, sheet_name="resources", dtype=str)
    except:
        return "File access error"

    resource_row = resources[resources["file"] == filename]

    if resource_row.empty:
        return "File not found"

    allowed = str(resource_row.iloc[0].get("allowed_users","")).split(",")
    blocked = str(resource_row.iloc[0].get("blocked_users","")).split(",")

    allowed = [a.strip() for a in allowed if a.strip()]
    blocked = [b.strip() for b in blocked if b.strip()]

    if not is_admin():
        if username in blocked:
            return "You are not allowed to download this file."

        if allowed and username not in allowed:
            return "You are not allowed to download this file."

    return send_from_directory(UPLOAD_FOLDER, filename)
    
    
    
    
    
@app.route("/favorite/<int:rid>")
def favorite(rid):

    if "user" not in session:
        return redirect("/login/user")

    username = session.get("user")

    df = pd.read_excel(EXCEL_FILE, sheet_name="resources", dtype=str)

    if "favorites" not in df.columns:
        df["favorites"] = ""

    # Get current favorites safely
    current = str(df.loc[df["id"].astype(int)==rid, "favorites"].values[0])

    # Convert to list (clean)
    users = [u.strip() for u in current.split(",") if u.strip()]

    # Toggle logic (add/remove properly)
    if username in users:
        users.remove(username)
    else:
        users.append(username)

    # Save back clean string
    df.loc[df["id"].astype(int)==rid, "favorites"] = ",".join(users)

    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name="resources", index=False)

    return redirect("/resources")
    
    
    
@app.route("/notifications")
def notifications():

    if "user" not in session:
        return redirect("/login/user")

    all_notes = load_json("notifications.json")

    user_notes = [
        n for n in all_notes
        if n.get("user") in [None, session["user"]]
    ]
    
    
    # 👇 ADD THIS
    session["notification_count"] = len(user_notes)
    # OPTIONAL RESET
    #session["notification_count"] = 0

    return render_template("notifications.html", notes=user_notes)
    
    
  #-------------------------------

@app.route("/search")
def search():
    return render_template("search.html")
    
    
@app.route("/logs")
def view_logs():

    if "user" not in session or not is_admin():
        return redirect("/login/admin")

    try:
        file_path = os.path.join(os.getcwd(), "users.xlsx")

        df = pd.read_excel(file_path, sheet_name="logs", engine="openpyxl")

        df = df.fillna("")
        df.columns = df.columns.str.strip().str.lower()

        logs = df.to_dict(orient="records")

    except Exception as e:
        print("ERROR:", e)
        logs = []

    return render_template("logs.html", logs=logs)


@app.route("/analytics")
def analytics():

    try:
        resources = pd.read_excel(EXCEL_FILE, sheet_name="resources", dtype=str)
        total_resources = len(resources)
    except:
        total_resources = 0

    try:
        help_requests = len(load_json("help_requests.json"))
    except:
        help_requests = 0

    try:
        users = pd.read_excel(EXCEL_FILE, sheet_name="users", dtype=str)
        users_count = len(users)
    except:
        users_count = 0

    return render_template(
        "analytics.html",
        total_resources=total_resources,
        help_requests=help_requests,
        users_count=users_count
    )

# -----------------------------
# Run App
# -----------------------------
if __name__ == "__main__":
    #app.run(debug=True)
    app.run(debug=True, host="0.0.0.0", port=5000)