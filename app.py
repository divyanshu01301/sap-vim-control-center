# ============================================================
# DPE Technologies – SAP VIM Control Center (Cloud Secure)
# Same functionality – Improved security
# ============================================================

from flask import Flask, render_template, request, redirect, session, jsonify, send_from_directory
from functools import wraps
import os
import math
import imaplib
from datetime import datetime
import vim_email_processor

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "supersecretkey")

# ============================================================
# CLOUD BASE CONFIG
# ============================================================

BASE_FOLDER = os.path.join(os.getcwd(), "data")

INCOMING_FOLDER = os.path.join(BASE_FOLDER, "incoming")
REJECTED_FOLDER = os.path.join(BASE_FOLDER, "rejected")
LOG_FOLDER = os.path.join(BASE_FOLDER, "logs")
LOG_FILE = os.path.join(LOG_FOLDER, "vim_log.log")

ITEMS_PER_PAGE = 5

for folder in [INCOMING_FOLDER, REJECTED_FOLDER, LOG_FOLDER]:
    os.makedirs(folder, exist_ok=True)


# ============================================================
# LOGIN REQUIRED DECORATOR
# ============================================================

def login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if "email" not in session:
            return redirect("/")
        return f(*args, **kwargs)
    return wrapper


# ============================================================
# LOGIN (WITH IMAP VALIDATION)
# ============================================================

@app.route("/", methods=["GET", "POST"])
def login():

    if request.method == "POST":
        email = request.form.get("email")
        password = request.form.get("password")

        if not email or not password:
            return render_template("login.html", error="Enter credentials")

        # Validate IMAP login before allowing dashboard
        try:
            mail = imaplib.IMAP4_SSL(os.environ.get("IMAP_SERVER", "imap.one.com"), 993)
            mail.login(email, password)
            mail.logout()
        except:
            return render_template("login.html", error="Invalid email or password")

        session["email"] = email
        session["password"] = password  # temporary only

        return redirect("/dashboard")

    return render_template("login.html")


# ============================================================
# DASHBOARD
# ============================================================

@app.route("/dashboard")
@login_required
def dashboard():

    message = session.pop("message", None)

    incoming_count = len(os.listdir(INCOMING_FOLDER))
    rejected_count = len(os.listdir(REJECTED_FOLDER))

    return render_template(
        "dashboard.html",
        user=session["email"],
        message=message,
        incoming_count=incoming_count,
        rejected_count=rejected_count
    )


# ============================================================
# START PROCESSING
# ============================================================

@app.route("/start")
@login_required
def start_processing():

    try:
        result = vim_email_processor.run_processor(
            session["email"],
            session["password"],
            INCOMING_FOLDER,
            REJECTED_FOLDER,
            LOG_FILE
        )

        session["message"] = result

        # Remove password after use (important)
        session.pop("password", None)

    except Exception as e:
        session["message"] = f"Error: {str(e)}"

    return redirect("/dashboard")


# ============================================================
# LIVE LOGS API
# ============================================================

@app.route("/api/logs")
@login_required
def get_logs():

    if not os.path.exists(LOG_FILE):
        return jsonify({"logs": []})

    with open(LOG_FILE, "r") as f:
        lines = f.readlines()[-40:]

    return jsonify({"logs": lines})


# ============================================================
# FOLDER HELPER
# ============================================================

def get_folder_data(folder_path, page):

    files = []

    for file in os.listdir(folder_path):
        full_path = os.path.join(folder_path, file)

        if os.path.isfile(full_path):
            files.append({
                "name": file,
                "size": round(os.path.getsize(full_path) / 1024, 2),
                "timestamp": datetime.fromtimestamp(
                    os.path.getmtime(full_path)
                ).strftime("%Y-%m-%d %H:%M:%S"),
                "modified_raw": os.path.getmtime(full_path)
            })

    files.sort(key=lambda x: x["modified_raw"], reverse=True)

    total_files = len(files)
    total_pages = max(1, math.ceil(total_files / ITEMS_PER_PAGE))

    start = (page - 1) * ITEMS_PER_PAGE
    end = start + ITEMS_PER_PAGE

    return files[start:end], total_pages, total_files


# ============================================================
# VIEW INCOMING
# ============================================================

@app.route("/incoming")
@login_required
def view_incoming():

    page = int(request.args.get("page", 1))

    files, total_pages, total_files = get_folder_data(INCOMING_FOLDER, page)

    return render_template(
        "folder_view.html",
        title="Incoming Documents",
        files=files,
        folder="incoming",
        page=page,
        total_pages=total_pages,
        total_files=total_files
    )


# ============================================================
# VIEW REJECTED
# ============================================================

@app.route("/rejected")
@login_required
def view_rejected():

    page = int(request.args.get("page", 1))

    files, total_pages, total_files = get_folder_data(REJECTED_FOLDER, page)

    return render_template(
        "folder_view.html",
        title="Rejected Documents",
        files=files,
        folder="rejected",
        page=page,
        total_pages=total_pages,
        total_files=total_files
    )


# ============================================================
# DOWNLOAD FILE
# ============================================================

@app.route("/file/<folder>/<filename>")
@login_required
def open_file(folder, filename):

    if ".." in filename:
        return "Invalid filename", 400

    if folder == "incoming":
        directory = INCOMING_FOLDER
    elif folder == "rejected":
        directory = REJECTED_FOLDER
    else:
        return "Invalid folder", 400

    return send_from_directory(directory, filename, as_attachment=True)


# ============================================================
# LOGOUT
# ============================================================

@app.route("/logout")
@login_required
def logout():
    session.clear()
    return redirect("/")


if __name__ == "__main__":
    app.run(debug=False)