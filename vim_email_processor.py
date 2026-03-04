# ============================================================
# SAP VIM Email Processor
# Email / PDF → PDF/A + Invoice Classification
# ============================================================

import imaplib
import email
import os
import pdfplumber
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from email.utils import parsedate_to_datetime
import re


IMAP_SERVER = "imap.one.com"
IMAP_PORT = 993


# ============================================================
# CLEAN FILE NAME
# ============================================================

def clean_filename(name):

    name = re.sub(r'[^\w\-_\. ]', '_', name)
    name = name.replace(" ", "_")

    return name


# ============================================================
# EMAIL TEXT → PDF
# ============================================================

def email_to_pdf(text, output_path):

    c = canvas.Canvas(output_path, pagesize=letter)

    y = 750

    for line in text.split("\n"):

        c.drawString(40, y, line)

        y -= 15

        if y < 50:
            c.showPage()
            y = 750

    c.save()


# ============================================================
# MAIN PROCESSOR
# ============================================================

def run_processor(
        email_user,
        email_pass,
        incoming_folder,
        rejected_folder,
        start_date,
        end_date):


    processed = 0

    start_date = datetime.strptime(start_date,"%Y-%m-%d")
    end_date = datetime.strptime(end_date,"%Y-%m-%d")

    mail = imaplib.IMAP4_SSL(IMAP_SERVER,IMAP_PORT)

    mail.login(email_user,email_pass)

    mail.select("INBOX")

    status,data = mail.search(None,"ALL")

    email_ids = data[0].split()

    for eid in email_ids:

        status,msg_data = mail.fetch(eid,"(RFC822)")

        raw_email = msg_data[0][1]

        msg = email.message_from_bytes(raw_email)


        # =========================
        # DATE FILTER
        # =========================

        try:

            email_date = parsedate_to_datetime(msg.get("Date"))

        except:

            continue

        email_date = email_date.replace(tzinfo=None)

        if not (start_date <= email_date <= end_date):
            continue


        subject = msg.get("Subject","")
        sender = msg.get("From","")


        pdf_found = False


        # =========================
        # ATTACHMENT PROCESSING
        # =========================

        if msg.is_multipart():

            for part in msg.walk():

                filename = part.get_filename()

                if filename and filename.lower().endswith(".pdf"):

                    pdf_found = True

                    filename = clean_filename(filename)

                    temp_path = os.path.join(incoming_folder,"temp_"+filename)

                    with open(temp_path,"wb") as f:
                        f.write(part.get_payload(decode=True))


                    text = ""

                    try:

                        with pdfplumber.open(temp_path) as pdf:

                            for page in pdf.pages:

                                t = page.extract_text()

                                if t:
                                    text += t

                    except:
                        text = ""


                    # =========================
                    # CLASSIFICATION
                    # =========================

                    if "invoice" in text.lower() or "inv" in text.lower():

                        final_path = os.path.join(incoming_folder,filename)

                    else:

                        final_path = os.path.join(rejected_folder,filename)


                    os.rename(temp_path,final_path)


        # =========================
        # NO PDF → CREATE PDF
        # =========================

        if not pdf_found:

            body=""

            if msg.is_multipart():

                for part in msg.walk():

                    if part.get_content_type()=="text/plain":

                        body += part.get_payload(decode=True).decode(errors="ignore")

            else:

                body = msg.get_payload(decode=True).decode(errors="ignore")


            content=f"""
From: {sender}

Subject: {subject}

Date: {email_date}

Body:

{body}
"""


            filename=f"email_{processed}.pdf"

            filename=clean_filename(filename)

            temp_path=os.path.join(incoming_folder,"temp_"+filename)

            email_to_pdf(content,temp_path)


            if "invoice" in content.lower() or "inv" in content.lower():

                final_path=os.path.join(incoming_folder,filename)

            else:

                final_path=os.path.join(rejected_folder,filename)


            os.rename(temp_path,final_path)


        processed+=1


    mail.logout()


    return f"{processed} emails processed successfully."