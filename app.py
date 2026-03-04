# ============================================================
# SAP VIM Control Center
# Email Processing with Date Filter + Preview + Download
# ============================================================

from flask import Flask, render_template, request, redirect, session, send_from_directory
from functools import wraps
from werkzeug.utils import secure_filename
import os
import vim_email_processor

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "supersecret")


# ============================================================
# CLOUD STORAGE PATHS
# ============================================================

BASE_FOLDER = os.path.join(os.getcwd(), "data")

INCOMING_FOLDER = os.path.join(BASE_FOLDER, "incoming")
REJECTED_FOLDER = os.path.join(BASE_FOLDER, "rejected")
LOG_FOLDER = os.path.join(BASE_FOLDER, "logs")

LOG_FILE = os.path.join(LOG_FOLDER, "vim_log.log")

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
# LOGIN PAGE
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
@login_required
def dashboard():

    message = session.pop("message", None)

    incoming_count = len(os.listdir(INCOMING_FOLDER))
    rejected_count = len(os.listdir(REJECTED_FOLDER))

    return render_template(
        "dashboard.html",
        user=session["email"],
        incoming_count=incoming_count,
        rejected_count=rejected_count,
        message=message
    )


# ============================================================
# PROCESS EMAILS WITH DATE RANGE
# ============================================================

@app.route("/process", methods=["POST"])
@login_required
def process_emails():

    start_date = request.form.get("start_date")
    end_date = request.form.get("end_date")

    try:

        result = vim_email_processor.run_processor(
            session["email"],
            session["password"],
            INCOMING_FOLDER,
            REJECTED_FOLDER,
            LOG_FILE,
            start_date,
            end_date
        )

        session["message"] = result

        # Remove password after processing (security)
        session.pop("password", None)

    except Exception as e:

        session["message"] = str(e)

    return redirect("/dashboard")


# ============================================================
# VIEW INCOMING DOCUMENTS
# ============================================================

@app.route("/incoming")
@login_required
def view_incoming():

    files = []

    for file in os.listdir(INCOMING_FOLDER):

        full_path = os.path.join(INCOMING_FOLDER, file)

        if os.path.isfile(full_path):

            files.append({
                "name": file,
                "size": round(os.path.getsize(full_path) / 1024, 2),
                "timestamp": os.path.getmtime(full_path)
            })

    return render_template(
        "folder_view.html",
        title="Incoming Documents",
        files=files,
        folder="incoming",
        total_files=len(files)
    )


# ============================================================
# VIEW REJECTED DOCUMENTS
# ============================================================

@app.route("/rejected")
@login_required
def view_rejected():

    files = []

    for file in os.listdir(REJECTED_FOLDER):

        full_path = os.path.join(REJECTED_FOLDER, file)

        if os.path.isfile(full_path):

            files.append({
                "name": file,
                "size": round(os.path.getsize(full_path) / 1024, 2),
                "timestamp": os.path.getmtime(full_path)
            })

    return render_template(
        "folder_view.html",
        title="Rejected Documents",
        files=files,
        folder="rejected",
        total_files=len(files)
    )


# ============================================================
# PREVIEW FILE
# ============================================================

@app.route("/file/<folder>/<filename>")
@login_required
def preview_file(folder, filename):

    filename = secure_filename(filename)

    if folder == "incoming":
        directory = INCOMING_FOLDER
    else:
        directory = REJECTED_FOLDER

    return send_from_directory(directory, filename)


# ============================================================
# DOWNLOAD FILE
# ============================================================

@app.route("/download/<folder>/<filename>")
@login_required
def download_file(folder, filename):

    filename = secure_filename(filename)

    if folder == "incoming":
        directory = INCOMING_FOLDER
    else:
        directory = REJECTED_FOLDER

    return send_from_directory(directory, filename, as_attachment=True)


# ============================================================
# LOGOUT
# ============================================================

@app.route("/logout")
def logout():

    session.clear()

    return redirect("/")


# ============================================================
# RUN LOCAL SERVER
# ============================================================

if __name__ == "__main__":
    app.run()