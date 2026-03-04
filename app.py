from flask import (Flask, render_template, request, redirect,
                   session, jsonify, send_from_directory, Response, stream_with_context)
from functools import wraps
from datetime import datetime, timedelta
import os
import threading
import time
import vim_email_processor

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "vim_dpe_2024_super_secret_key_x9z")
app.config["SESSION_COOKIE_SAMESITE"]    = "Lax"
app.config["SESSION_COOKIE_SECURE"]      = False
app.config["SESSION_COOKIE_HTTPONLY"]    = True
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(hours=24)

BASE_FOLDER     = os.path.join(os.getcwd(), "data")
INCOMING_FOLDER = os.path.join(BASE_FOLDER, "incoming")
REJECTED_FOLDER = os.path.join(BASE_FOLDER, "rejected")
LOG_FOLDER      = os.path.join(BASE_FOLDER, "logs")
LOG_FILE        = os.path.join(LOG_FOLDER,  "vim_log.log")

for folder in [INCOMING_FOLDER, REJECTED_FOLDER, LOG_FOLDER]:
    os.makedirs(folder, exist_ok=True)

_processing_lock  = threading.Lock()
processing_status = {"running": False, "result": None}


def login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if "email" not in session:
            return redirect("/")
        return f(*args, **kwargs)
    return wrapper


@app.route("/", methods=["GET", "POST"])
def login():
    if "email" in session:
        return redirect("/dashboard")
    if request.method == "POST":
        email_val = request.form.get("email", "").strip()
        password  = request.form.get("password", "")
        if not email_val or not password:
            return render_template("login.html", error="Please enter both email and password.")
        session.permanent = True
        session["email"]    = email_val
        session["password"] = password
        return redirect("/dashboard")
    return render_template("login.html", error=None)


@app.route("/dashboard")
@login_required
def dashboard():
    incoming_count = len([f for f in os.listdir(INCOMING_FOLDER)
                          if os.path.isfile(os.path.join(INCOMING_FOLDER, f))])
    rejected_count = len([f for f in os.listdir(REJECTED_FOLDER)
                          if os.path.isfile(os.path.join(REJECTED_FOLDER, f))])
    message = session.pop("message", None)
    return render_template("dashboard.html",
                           user=session["email"],
                           incoming_count=incoming_count,
                           rejected_count=rejected_count,
                           message=message)


def _run_in_background(email_user, email_pass, start_date, end_date):
    global processing_status
    try:
        result = vim_email_processor.run_processor(
            email_user=email_user,
            email_pass=email_pass,
            incoming_folder=INCOMING_FOLDER,
            rejected_folder=REJECTED_FOLDER,
            log_file=LOG_FILE,
            start_date=start_date,
            end_date=end_date,
        )
        processing_status["result"] = result
    except Exception as e:
        processing_status["result"] = f"Error: {str(e)}"
    finally:
        processing_status["running"] = False


@app.route("/process", methods=["POST"])
@login_required
def process_emails():
    global processing_status
    with _processing_lock:
        if processing_status["running"]:
            session["message"] = "Processing already running. Watch the live log."
            return redirect("/dashboard")
        start_date = request.form.get("start_date", "").strip()
        end_date   = request.form.get("end_date",   "").strip()
        if not start_date or not end_date:
            session["message"] = "Please select both a start and end date."
            return redirect("/dashboard")
        processing_status["running"] = True
        processing_status["result"]  = None
        t = threading.Thread(
            target=_run_in_background,
            args=(session["email"], session["password"], start_date, end_date),
            daemon=True,
        )
        t.start()
    session["message"] = f"Processing started for {start_date} to {end_date}."
    return redirect("/dashboard")


@app.route("/api/status")
@login_required
def api_status():
    return jsonify(processing_status)


@app.route("/api/counts")
@login_required
def api_counts():
    inc = len([f for f in os.listdir(INCOMING_FOLDER)
               if os.path.isfile(os.path.join(INCOMING_FOLDER, f))])
    rej = len([f for f in os.listdir(REJECTED_FOLDER)
               if os.path.isfile(os.path.join(REJECTED_FOLDER, f))])
    return jsonify({"incoming": inc, "rejected": rej})


@app.route("/api/log-stream")
@login_required
def log_stream():
    def generate():
        last_pos   = 0
        idle_ticks = 0
        while True:
            if os.path.exists(LOG_FILE):
                try:
                    with open(LOG_FILE, "r", encoding="utf-8", errors="ignore") as f:
                        f.seek(last_pos)
                        chunk    = f.read()
                        last_pos = f.tell()
                    if chunk:
                        idle_ticks = 0
                        for line in chunk.splitlines():
                            line = line.strip()
                            if line:
                                yield f"data: {line}\n\n"
                    else:
                        idle_ticks += 1
                        if idle_ticks % 20 == 0:
                            yield ": keep-alive\n\n"
                except Exception:
                    pass
            else:
                yield "data: Waiting for log file...\n\n"
            time.sleep(0.5)
    return Response(
        stream_with_context(generate()),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


def _list_files(folder):
    files = []
    for f in sorted(os.listdir(folder)):
        full = os.path.join(folder, f)
        if os.path.isfile(full):
            files.append({
                "name":      f,
                "size":      round(os.path.getsize(full) / 1024, 2),
                "timestamp": datetime.fromtimestamp(
                    os.path.getmtime(full)).strftime("%d %b %Y %H:%M"),
            })
    return files


@app.route("/incoming")
@login_required
def incoming():
    files = _list_files(INCOMING_FOLDER)
    return render_template("folder_view.html",
                           title="Incoming Invoices", files=files,
                           folder="incoming", total_files=len(files))


@app.route("/rejected")
@login_required
def rejected():
    files = _list_files(REJECTED_FOLDER)
    return render_template("folder_view.html",
                           title="Rejected Documents", files=files,
                           folder="rejected", total_files=len(files))


@app.route("/file/<folder>/<filename>")
@login_required
def preview(folder, filename):
    directory = INCOMING_FOLDER if folder == "incoming" else REJECTED_FOLDER
    return send_from_directory(directory, filename)


@app.route("/download/<folder>/<filename>")
@login_required
def download(folder, filename):
    directory = INCOMING_FOLDER if folder == "incoming" else REJECTED_FOLDER
    return send_from_directory(directory, filename, as_attachment=True)


@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False)