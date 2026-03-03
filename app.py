# ============================================================
# DPE Technologies – SAP VIM Control Center (Full Version)
# Login + Dashboard + Live Logs + Folder Viewer
# ============================================================

from flask import Flask, render_template, request, redirect, session, jsonify, send_from_directory
import os
import math
from datetime import datetime
import vim_email_processor

app = Flask(__name__)
app.secret_key = os.urandom(32)

# ============================================================
# BASE CONFIG
# ============================================================

BASE_FOLDER = r"C:\VIM_AUTOMATION"
INCOMING_FOLDER = os.path.join(BASE_FOLDER, "incoming")
REJECTED_FOLDER = os.path.join(BASE_FOLDER, "rejected")
LOG_FILE = os.path.join(BASE_FOLDER, "logs", "vim_log.log")

ITEMS_PER_PAGE = 5


# ============================================================
# LOGIN
# ============================================================

@app.route("/", methods=["GET", "POST"])
def login():

    if request.method == "POST":
        email = request.form.get("email")
        password = request.form.get("password")

        if not email or not password:
            return render_template("login.html", error="Enter credentials")

        session["email"] = email
        session["password"] = password

        return redirect("/dashboard")

    return render_template("login.html")


# ============================================================
# DASHBOARD
# ============================================================

@app.route("/dashboard")
def dashboard():

    if "email" not in session:
        return redirect("/")

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
def start_processing():

    if "email" not in session:
        return redirect("/")

    try:
        result = vim_email_processor.run_processor(
            session["email"],
            session["password"]
        )
        session["message"] = result

    except Exception as e:
        session["message"] = f"Error: {str(e)}"

    return redirect("/dashboard")


# ============================================================
# LIVE LOGS API
# ============================================================

@app.route("/api/logs")
def get_logs():

    if not os.path.exists(LOG_FILE):
        return jsonify({"logs": []})

    with open(LOG_FILE, "r") as f:
        lines = f.readlines()[-40:]

    return jsonify({"logs": lines})


# ============================================================
# FOLDER DATA HELPER
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

    # Sort latest first
    files.sort(key=lambda x: x["modified_raw"], reverse=True)

    total_files = len(files)
    total_pages = max(1, math.ceil(total_files / ITEMS_PER_PAGE))

    start = (page - 1) * ITEMS_PER_PAGE
    end = start + ITEMS_PER_PAGE

    paginated_files = files[start:end]

    return paginated_files, total_pages, total_files


# ============================================================
# VIEW INCOMING
# ============================================================

@app.route("/incoming")
def view_incoming():

    if "email" not in session:
        return redirect("/")

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
def view_rejected():

    if "email" not in session:
        return redirect("/")

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
# OPEN / DOWNLOAD FILE
# ============================================================

@app.route("/file/<folder>/<filename>")
def open_file(folder, filename):

    if folder == "incoming":
        directory = INCOMING_FOLDER
    else:
        directory = REJECTED_FOLDER

    return send_from_directory(directory, filename)


# ============================================================
# LOGOUT
# ============================================================

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")


# ============================================================
# RUN SERVER
# ============================================================

if __name__ == "__main__":
    app.run()