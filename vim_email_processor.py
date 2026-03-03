# ============================================================
# SAP VIM Email Processor – Cloud Production Version
# Dynamic Credentials + Render Compatible
# ============================================================

import imaplib
import email
import os
import pdfplumber
import logging
from datetime import datetime


# ============================================================
# CONFIGURATION
# ============================================================

IMAP_SERVER = "imap.one.com"
IMAP_PORT = 993


# ============================================================
# MAIN FUNCTION (MATCHES app.py)
# ============================================================

def run_processor(email_user, email_pass, incoming_folder, rejected_folder, log_file):

    # Setup logging dynamically (important for cloud)
    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s"
    )

    print(f"Starting processing for {email_user}")
    logging.info(f"Processing started for {email_user}")

    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(email_user, email_pass)
        mail.select("INBOX")

        logging.info("Mailbox connected")

    except Exception as e:
        logging.error("Login failed: " + str(e))
        return "Login Failed"

    status, messages = mail.search(None, 'UNSEEN')
    email_ids = messages[0].split()

    logging.info(f"Found {len(email_ids)} unread emails")

    for email_id in email_ids:

        try:
            status, msg_data = mail.fetch(email_id, "(RFC822)")
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)

            subject = msg.get("Subject", "")
            sender = msg.get("From", "")

            logging.info(f"Processing email from {sender}")

            if msg.is_multipart():

                for part in msg.walk():

                    filename = part.get_filename()

                    if filename and filename.lower().endswith(".pdf"):

                        temp_path = os.path.join(incoming_folder, "temp_" + filename)

                        with open(temp_path, "wb") as f:
                            f.write(part.get_payload(decode=True))

                        # Extract text
                        pdf_text = ""

                        try:
                            with pdfplumber.open(temp_path) as pdf:
                                for page in pdf.pages:
                                    text = page.extract_text()
                                    if text:
                                        pdf_text += text
                        except Exception as e:
                            logging.error("PDF read error: " + str(e))

                        # Classification Logic
                        classification = (
                            "INCOMING"
                            if "invoice" in pdf_text.lower()
                            else "REJECTED"
                        )

                        destination_folder = (
                            incoming_folder
                            if classification == "INCOMING"
                            else rejected_folder
                        )

                        final_path = os.path.join(destination_folder, filename)

                        os.rename(temp_path, final_path)

                        logging.info(f"{filename} moved to {classification}")

            mail.store(email_id, '+FLAGS', '\\Seen')

        except Exception as e:
            logging.error(str(e))

    mail.logout()

    logging.info("Processing completed")

    return "Processing Completed Successfully"