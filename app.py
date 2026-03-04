from flask import Flask, render_template, request, redirect, session, send_from_directory
from functools import wraps
from datetime import datetime
import os
import vim_email_processor

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "vim_control_center")

BASE_FOLDER = os.path.join(os.getcwd(), "data")
os.makedirs(BASE_FOLDER, exist_ok=True)


# =========================================================
# USER SPECIFIC FOLDERS
# =========================================================

def get_user_folders():

    user = session["email"].replace("@","_").replace(".","_")

    user_folder = os.path.join(BASE_FOLDER, user)

    incoming = os.path.join(user_folder,"incoming")
    rejected = os.path.join(user_folder,"rejected")

    os.makedirs(incoming, exist_ok=True)
    os.makedirs(rejected, exist_ok=True)

    return incoming, rejected


# =========================================================
# LOGIN REQUIRED DECORATOR
# =========================================================

def login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if "email" not in session:
            return redirect("/")
        return f(*args, **kwargs)
    return wrapper


# =========================================================
# LOGIN PAGE
# =========================================================

@app.route("/", methods=["GET","POST"])
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


# =========================================================
# DASHBOARD
# =========================================================

@app.route("/dashboard")
@login_required
def dashboard():

    incoming_folder, rejected_folder = get_user_folders()

    incoming_count = len(os.listdir(incoming_folder))
    rejected_count = len(os.listdir(rejected_folder))

    message = session.pop("message", None)

    return render_template(
        "dashboard.html",
        user=session["email"],
        incoming_count=incoming_count,
        rejected_count=rejected_count,
        message=message
    )


# =========================================================
# PROCESS EMAILS
# =========================================================

@app.route("/process", methods=["POST"])
@login_required
def process():

    start_date = request.form.get("start_date")
    end_date = request.form.get("end_date")

    incoming_folder, rejected_folder = get_user_folders()

    try:

        result = vim_email_processor.run_processor(
            session["email"],
            session["password"],
            incoming_folder,
            rejected_folder,
            start_date,
            end_date
        )

        session["message"] = result

    except Exception as e:

        import traceback
        print(traceback.format_exc())

        session["message"] = f"Processing failed: {str(e)}"

    return redirect("/dashboard")


# =========================================================
# INCOMING DOCUMENTS
# =========================================================

@app.route("/incoming")
@login_required
def incoming():

    incoming_folder, rejected_folder = get_user_folders()

    files = []

    for f in os.listdir(incoming_folder):

        path = os.path.join(incoming_folder, f)

        if os.path.isfile(path):

            files.append({
                "name": f,
                "size": round(os.path.getsize(path) / 1024, 2),
                "timestamp": datetime.fromtimestamp(
                    os.path.getmtime(path)
                ).strftime("%Y-%m-%d %H:%M")
            })

    return render_template(
        "folder_view.html",
        title="Incoming Documents",
        files=files,
        folder="incoming"
    )


# =========================================================
# REJECTED DOCUMENTS
# =========================================================

@app.route("/rejected")
@login_required
def rejected():

    incoming_folder, rejected_folder = get_user_folders()

    files = []

    for f in os.listdir(rejected_folder):

        path = os.path.join(rejected_folder, f)

        if os.path.isfile(path):

            files.append({
                "name": f,
                "size": round(os.path.getsize(path) / 1024, 2),
                "timestamp": datetime.fromtimestamp(
                    os.path.getmtime(path)
                ).strftime("%Y-%m-%d %H:%M")
            })

    return render_template(
        "folder_view.html",
        title="Rejected Documents",
        files=files,
        folder="rejected"
    )


# =========================================================
# PREVIEW FILE
# =========================================================

@app.route("/file/<folder>/<filename>")
@login_required
def preview(folder, filename):

    incoming_folder, rejected_folder = get_user_folders()

    if folder == "incoming":
        directory = incoming_folder
    elif folder == "rejected":
        directory = rejected_folder
    else:
        return "Invalid folder", 400

    return send_from_directory(directory, filename)


# =========================================================
# DOWNLOAD FILE
# =========================================================

@app.route("/download/<folder>/<filename>")
@login_required
def download(folder, filename):

    incoming_folder, rejected_folder = get_user_folders()

    if folder == "incoming":
        directory = incoming_folder
    elif folder == "rejected":
        directory = rejected_folder
    else:
        return "Invalid folder", 400

    return send_from_directory(directory, filename, as_attachment=True)


# =========================================================
# LOGOUT
# =========================================================

@app.route("/logout")
def logout():

    session.clear()

    return redirect("/")


# =========================================================
# RUN SERVER
# =========================================================

if __name__ == "__main__":

    port = int(os.environ.get("PORT", 10000))

    app.run(host="0.0.0.0", port=port)