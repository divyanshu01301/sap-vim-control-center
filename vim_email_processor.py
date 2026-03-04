# ============================================================
# SAP VIM Email Processor
# PDF / Email → PDF/A + Invoice Classification
# ============================================================

import imaplib
import email
import os
import pdfplumber
import logging
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas


IMAP_SERVER = "imap.one.com"
IMAP_PORT = 993


def convert_email_to_pdf(content, output_path):

    c = canvas.Canvas(output_path, pagesize=letter)

    text = c.beginText(40, 750)
    text.setFont("Helvetica", 10)

    for line in content.split("\n"):
        text.textLine(line)

    c.drawText(text)
    c.save()


# ============================================================
# MAIN PROCESSOR
# ============================================================

def run_processor(email_user, email_pass,
                  incoming_folder,
                  rejected_folder,
                  log_file,
                  start_date,
                  end_date):

    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s"
    )

    processed_count = 0

    try:

        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(email_user, email_pass)
        mail.select("INBOX")

    except Exception as e:

        return f"Login failed: {str(e)}"

    status, messages = mail.search(None, "ALL")
    email_ids = messages[0].split()

    start_date = datetime.strptime(start_date, "%Y-%m-%d")
    end_date = datetime.strptime(end_date, "%Y-%m-%d")

    for email_id in email_ids:

        status, msg_data = mail.fetch(email_id, "(RFC822)")
        raw_email = msg_data[0][1]

        msg = email.message_from_bytes(raw_email)

        email_date = msg.get("Date")

        try:
            email_datetime = email.utils.parsedate_to_datetime(email_date)
        except:
            continue

        if not (start_date <= email_datetime.replace(tzinfo=None) <= end_date):
            continue

        subject = msg.get("Subject", "")
        sender = msg.get("From", "")

        body_text = ""

        pdf_found = False

        if msg.is_multipart():

            for part in msg.walk():

                filename = part.get_filename()

                if filename and filename.lower().endswith(".pdf"):

                    pdf_found = True

                    filepath = os.path.join(incoming_folder, filename)

                    with open(filepath, "wb") as f:
                        f.write(part.get_payload(decode=True))

                    try:

                        text = ""

                        with pdfplumber.open(filepath) as pdf:

                            for page in pdf.pages:
                                page_text = page.extract_text()
                                if page_text:
                                    text += page_text

                        if "invoice" in text.lower() or "inv" in text.lower():

                            os.rename(
                                filepath,
                                os.path.join(incoming_folder, filename)
                            )

                        else:

                            os.rename(
                                filepath,
                                os.path.join(rejected_folder, filename)
                            )

                    except:

                        os.rename(
                            filepath,
                            os.path.join(rejected_folder, filename)
                        )

        if not pdf_found:

            if msg.is_multipart():

                for part in msg.walk():

                    if part.get_content_type() == "text/plain":

                        body_text += part.get_payload(decode=True).decode(errors="ignore")

            else:

                body_text = msg.get_payload(decode=True).decode(errors="ignore")

            full_text = f"""
From: {sender}

Subject: {subject}

Date: {email_date}

Body:

{body_text}
"""

            filename = f"email_{processed_count}.pdf"

            pdf_path = os.path.join(incoming_folder, filename)

            convert_email_to_pdf(full_text, pdf_path)

            if "invoice" in full_text.lower() or "inv" in full_text.lower():

                os.rename(pdf_path, os.path.join(incoming_folder, filename))

            else:

                os.rename(pdf_path, os.path.join(rejected_folder, filename))

        processed_count += 1

    mail.logout()

    return f"Processing completed. {processed_count} emails processed."