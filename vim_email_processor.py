# ============================================================
# SAP VIM Email Processor – Dynamic Enterprise Version
# No Hardcoded Credentials
# ============================================================

import imaplib
import email
import os
import pdfplumber
import logging
import subprocess
import sqlite3
from datetime import datetime

# ============================================================
# CONFIGURATION
# ============================================================

IMAP_SERVER = "imap.one.com"
IMAP_PORT = 993

BASE_FOLDER = r"C:\VIM_AUTOMATION"

INCOMING_FOLDER = os.path.join(BASE_FOLDER, "incoming")
PROCESSED_FOLDER = os.path.join(BASE_FOLDER, "processed")
REJECTED_FOLDER = os.path.join(BASE_FOLDER, "rejected")
LOG_FOLDER = os.path.join(BASE_FOLDER, "logs")

LOG_FILE = os.path.join(LOG_FOLDER, "vim_log.log")

# Create folders if not exist
for folder in [INCOMING_FOLDER, PROCESSED_FOLDER, REJECTED_FOLDER, LOG_FOLDER]:
    os.makedirs(folder, exist_ok=True)

# Logging
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# ============================================================
# MAIN FUNCTION (DYNAMIC LOGIN)
# ============================================================

def run_processor(email_user, email_pass):

    print(f"Starting processing for {email_user}")
    logging.info(f"Processing started for {email_user}")

    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(email_user, email_pass)
        mail.select("INBOX")
        print("Connected to mailbox")
        logging.info("Mailbox connected")

    except Exception as e:
        logging.error("Login failed: " + str(e))
        return "Login Failed"

    status, messages = mail.search(None, 'UNSEEN')
    email_ids = messages[0].split()

    print(f"Found {len(email_ids)} unread emails")

    for email_id in email_ids:

        try:
            status, msg_data = mail.fetch(email_id, "(RFC822)")
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)

            subject = msg.get("Subject", "")
            sender = msg.get("From", "")

            print(f"Processing email from {sender}")
            logging.info(f"Processing email from {sender}")

            if msg.is_multipart():

                for part in msg.walk():

                    filename = part.get_filename()

                    if filename and filename.lower().endswith(".pdf"):

                        temp_path = os.path.join(PROCESSED_FOLDER, filename)

                        with open(temp_path, "wb") as f:
                            f.write(part.get_payload(decode=True))

                        # Extract text for classification
                        pdf_text = ""
                        try:
                            with pdfplumber.open(temp_path) as pdf:
                                for page in pdf.pages:
                                    text = page.extract_text()
                                    if text:
                                        pdf_text += text
                        except:
                            pass

                        classification = "INCOMING" if "invoice" in pdf_text.lower() else "REJECTED"

                        destination = os.path.join(
                            INCOMING_FOLDER if classification == "INCOMING" else REJECTED_FOLDER,
                            filename
                        )

                        os.rename(temp_path, destination)

                        logging.info(f"{filename} moved to {classification}")

            mail.store(email_id, '+FLAGS', '\\Seen')

        except Exception as e:
            logging.error(str(e))

    mail.logout()

    logging.info("Processing completed")
    print("Processing complete")

    return "Processing Completed Successfully"