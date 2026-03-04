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


# ============================================================
# EMAIL → PDF
# ============================================================

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

    logging.info("===== VIM Email Processing Started =====")

    processed_count = 0

    try:

        logging.info("Connecting to mailbox")

        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(email_user, email_pass)
        mail.select("INBOX")

        logging.info("Mailbox connected successfully")

    except Exception as e:

        logging.error(f"Login failed: {str(e)}")
        return f"Login failed: {str(e)}"

    status, messages = mail.search(None, "ALL")
    email_ids = messages[0].split()

    logging.info(f"{len(email_ids)} emails found in mailbox")

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
            logging.warning("Failed to parse email date")
            continue

        if not (start_date <= email_datetime.replace(tzinfo=None) <= end_date):
            continue

        subject = msg.get("Subject", "")
        sender = msg.get("From", "")

        logging.info(f"Processing email from {sender} | Subject: {subject}")

        body_text = ""

        pdf_found = False

        if msg.is_multipart():

            for part in msg.walk():

                filename = part.get_filename()

                if filename and filename.lower().endswith(".pdf"):

                    pdf_found = True

                    filepath = os.path.join(incoming_folder, filename)

                    logging.info(f"PDF attachment detected: {filename}")

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

                            logging.info(f"Invoice detected in PDF: {filename}")

                            os.rename(
                                filepath,
                                os.path.join(incoming_folder, filename)
                            )

                        else:

                            logging.info(f"Rejected PDF (not invoice): {filename}")

                            os.rename(
                                filepath,
                                os.path.join(rejected_folder, filename)
                            )

                    except Exception as e:

                        logging.error(f"PDF parsing failed: {str(e)}")

                        os.rename(
                            filepath,
                            os.path.join(rejected_folder, filename)
                        )

        if not pdf_found:

            logging.info("No PDF attachment found — converting email to PDF")

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

                logging.info("Invoice keyword detected in email body")

                os.rename(pdf_path, os.path.join(incoming_folder, filename))

            else:

                logging.info("Email rejected (not invoice)")

                os.rename(pdf_path, os.path.join(rejected_folder, filename))

        processed_count += 1

    mail.logout()

    logging.info(f"Processing finished. {processed_count} emails processed.")

    return f"Processing completed. {processed_count} emails processed."