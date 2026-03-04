from flask import Flask, render_template, request, redirect, session, send_from_directory
from functools import wraps
from datetime import datetime
import os
import vim_email_processor

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY","vim_control")

BASE_FOLDER = os.path.join(os.getcwd(),"data")

INCOMING_FOLDER = os.path.join(BASE_FOLDER,"incoming")
REJECTED_FOLDER = os.path.join(BASE_FOLDER,"rejected")

for folder in [INCOMING_FOLDER,REJECTED_FOLDER]:
    os.makedirs(folder,exist_ok=True)

# ==========================
# LOGIN REQUIRED
# ==========================

def login_required(f):
    @wraps(f)
    def wrapper(*args,**kwargs):
        if "email" not in session:
            return redirect("/")
        return f(*args,**kwargs)
    return wrapper


# ==========================
# LOGIN
# ==========================

@app.route("/",methods=["GET","POST"])
def login():

    if request.method == "POST":

        email = request.form.get("email")
        password = request.form.get("password")

        if not email or not password:
            return render_template("login.html",error="Enter credentials")

        session["email"] = email
        session["password"] = password

        return redirect("/dashboard")

    return render_template("login.html")


# ==========================
# DASHBOARD
# ==========================

@app.route("/dashboard")
@login_required
def dashboard():

    incoming_count = len(os.listdir(INCOMING_FOLDER))
    rejected_count = len(os.listdir(REJECTED_FOLDER))

    message = session.pop("message",None)

    return render_template(
        "dashboard.html",
        user=session["email"],
        incoming_count=incoming_count,
        rejected_count=rejected_count,
        message=message
    )


# ==========================
# PROCESS EMAILS
# ==========================

@app.route("/process",methods=["POST"])
@login_required
def process():

    start_date = request.form.get("start_date")
    end_date = request.form.get("end_date")

    try:

        result = vim_email_processor.run_processor(
            session["email"],
            session["password"],
            INCOMING_FOLDER,
            REJECTED_FOLDER,
            start_date,
            end_date
        )

        session["message"] = result

    except Exception as e:

        session["message"] = str(e)

    return redirect("/dashboard")


# ==========================
# INCOMING
# ==========================

@app.route("/incoming")
@login_required
def incoming():

    files=[]

    for f in os.listdir(INCOMING_FOLDER):

        path=os.path.join(INCOMING_FOLDER,f)

        if os.path.isfile(path):

            files.append({
                "name":f,
                "size":round(os.path.getsize(path)/1024,2),
                "timestamp":datetime.fromtimestamp(
                    os.path.getmtime(path)
                ).strftime("%Y-%m-%d %H:%M")
            })

    return render_template(
        "folder_view.html",
        title="Incoming Documents",
        files=files,
        folder="incoming"
    )


# ==========================
# REJECTED
# ==========================

@app.route("/rejected")
@login_required
def rejected():

    files=[]

    for f in os.listdir(REJECTED_FOLDER):

        path=os.path.join(REJECTED_FOLDER,f)

        if os.path.isfile(path):

            files.append({
                "name":f,
                "size":round(os.path.getsize(path)/1024,2),
                "timestamp":datetime.fromtimestamp(
                    os.path.getmtime(path)
                ).strftime("%Y-%m-%d %H:%M")
            })

    return render_template(
        "folder_view.html",
        title="Rejected Documents",
        files=files,
        folder="rejected"
    )


# ==========================
# PREVIEW PDF
# ==========================

@app.route("/file/<folder>/<filename>")
@login_required
def preview(folder,filename):

    directory = INCOMING_FOLDER if folder=="incoming" else REJECTED_FOLDER

    return send_from_directory(directory,filename)


# ==========================
# DOWNLOAD
# ==========================

@app.route("/download/<folder>/<filename>")
@login_required
def download(folder,filename):

    directory = INCOMING_FOLDER if folder=="incoming" else REJECTED_FOLDER

    return send_from_directory(directory,filename,as_attachment=True)


# ==========================
# LOGOUT
# ==========================

@app.route("/logout")
def logout():

    session.clear()
    return redirect("/")


if __name__ == "__main__":
    app.run()