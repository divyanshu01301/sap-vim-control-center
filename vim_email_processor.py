# ============================================================
# SAP VIM Email Processor
# PDF + Email → PDF/A Conversion
# Invoice Classification
# ============================================================

import imaplib
import email
import os
import logging
import pdfplumber
import subprocess
from email.header import decode_header
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

IMAP_SERVER = "imap.one.com"
IMAP_PORT = 993


# ============================================================
# PDF → PDF/A Conversion
# ============================================================

def convert_to_pdfa(input_pdf, output_pdf):

    try:

        subprocess.run([
            "gs",
            "-dPDFA=2",
            "-dBATCH",
            "-dNOPAUSE",
            "-dNOOUTERSAVE",
            "-sDEVICE=pdfwrite",
            "-sOutputFile=" + output_pdf,
            input_pdf
        ])

    except Exception as e:
        print("PDF/A conversion failed:", e)


# ============================================================
# Email → PDF
# ============================================================

def email_to_pdf(email_data, output_pdf):

    c = canvas.Canvas(output_pdf, pagesize=letter)

    text = c.beginText(40, 750)
    text.setFont("Helvetica", 11)

    lines = email_data.split("\n")

    for line in lines:
        text.textLine(line)

    c.drawText(text)
    c.save()


# ============================================================
# Extract Email Body
# ============================================================

def get_email_body(msg):

    body = ""

    if msg.is_multipart():

        for part in msg.walk():

            content_type = part.get_content_type()

            if content_type == "text/plain":

                try:
                    body = part.get_payload(decode=True).decode()
                except:
                    pass

    else:

        try:
            body = msg.get_payload(decode=True).decode()
        except:
            pass

    return body


# ============================================================
# Classification Rule
# ============================================================

def classify_document(text):

    text = text.lower()

    if "invoice" in text or "inv" in text:
        return "INCOMING"

    return "REJECTED"


# ============================================================
# Main Processor
# ============================================================

def run_processor(email_user, email_pass, incoming_folder, rejected_folder, log_file):

    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s"
    )

    logging.info("Starting email processing")

    try:

        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(email_user, email_pass)
        mail.select("INBOX")

    except Exception as e:

        logging.error("Mailbox login failed: " + str(e))
        return "Login Failed"

    status, messages = mail.search(None, 'UNSEEN')

    email_ids = messages[0].split()

    logging.info(f"{len(email_ids)} new emails found")

    for email_id in email_ids:

        try:

            status, msg_data = mail.fetch(email_id, "(RFC822)")

            raw_email = msg_data[0][1]

            msg = email.message_from_bytes(raw_email)

            subject = msg.get("Subject", "")
            sender = msg.get("From", "")
            date = msg.get("Date", "")

            body = get_email_body(msg)

            full_text = subject + " " + body

            classification = classify_document(full_text)

            destination_folder = incoming_folder if classification == "INCOMING" else rejected_folder

            attachment_found = False

            # =================================================
            # Check Attachments
            # =================================================

            for part in msg.walk():

                filename = part.get_filename()

                if filename:

                    decoded_filename = decode_header(filename)[0][0]

                    if isinstance(decoded_filename, bytes):
                        decoded_filename = decoded_filename.decode()

                    if decoded_filename.lower().endswith(".pdf"):

                        attachment_found = True

                        temp_path = os.path.join(destination_folder, "temp_" + decoded_filename)

                        final_path = os.path.join(destination_folder, decoded_filename)

                        with open(temp_path, "wb") as f:
                            f.write(part.get_payload(decode=True))

                        # Convert to PDF/A
                        convert_to_pdfa(temp_path, final_path)

                        os.remove(temp_path)

                        logging.info(f"PDF processed: {decoded_filename}")

            # =================================================
            # If NO PDF attachment
            # =================================================

            if not attachment_found:

                email_content = f"""
From: {sender}
Date: {date}
Subject: {subject}

Body:
{body}
"""

                temp_pdf = os.path.join(destination_folder, "email_temp.pdf")
                final_pdf = os.path.join(destination_folder, "email_document.pdf")

                email_to_pdf(email_content, temp_pdf)

                convert_to_pdfa(temp_pdf, final_pdf)

                os.remove(temp_pdf)

                logging.info("Email converted to PDF/A")

            mail.store(email_id, '+FLAGS', '\\Seen')

        except Exception as e:

            logging.error("Processing error: " + str(e))

    mail.logout()

    logging.info("Processing completed")

    return "Processing Completed Successfully"