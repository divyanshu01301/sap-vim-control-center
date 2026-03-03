# ============================================================
# SAP VIM Email Processor – Cloud Production Version
# Clean Filenames + No UUID Garbage Names
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
# MAIN FUNCTION
# ============================================================

def run_processor(email_user, email_pass, incoming_folder, rejected_folder, log_file):

    # Ensure folders exist
    os.makedirs(incoming_folder, exist_ok=True)
    os.makedirs(rejected_folder, exist_ok=True)
    os.makedirs(os.path.dirname(log_file), exist_ok=True)

    # Setup logging
    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s"
    )

    logging.info("==========================================")
    logging.info(f"Processing started for {email_user}")
    logging.info("==========================================")

    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(email_user, email_pass)
        mail.select("INBOX")
        logging.info("Mailbox connected")

    except Exception as e:
        logging.error("Login failed: " + str(e))
        return "Login Failed"

    try:
        status, messages = mail.search(None, 'UNSEEN')
        email_ids = messages[0].split()
        logging.info(f"Found {len(email_ids)} unread emails")

    except Exception as e:
        logging.error("Search failed: " + str(e))
        return "Failed to read inbox"

    for email_id in email_ids:

        try:
            status, msg_data = mail.fetch(email_id, "(RFC822)")
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)

            subject = msg.get("Subject", "")
            sender = msg.get("From", "")

            logging.info(f"Processing email from: {sender}")
            logging.info(f"Subject: {subject}")

            if msg.is_multipart():

                for part in msg.walk():

                    filename = part.get_filename()

                    if filename and filename.lower().endswith(".pdf"):

                        # ============================================
                        # CLEAN FILENAME
                        # ============================================

                        clean_name = (
                            filename.replace(" ", "_")
                                    .replace("(", "")
                                    .replace(")", "")
                        )

                        final_name = clean_name
                        counter = 1

                        # Prevent duplicate files
                        while os.path.exists(os.path.join(incoming_folder, final_name)) or \
                              os.path.exists(os.path.join(rejected_folder, final_name)):

                            name, ext = os.path.splitext(clean_name)
                            final_name = f"{name}_{counter}{ext}"
                            counter += 1

                        temp_path = os.path.join(incoming_folder, final_name)

                        # Save PDF
                        with open(temp_path, "wb") as f:
                            f.write(part.get_payload(decode=True))

                        logging.info(f"Saved attachment: {final_name}")

                        # ============================================
                        # EXTRACT TEXT
                        # ============================================

                        pdf_text = ""

                        try:
                            with pdfplumber.open(temp_path) as pdf:
                                for page in pdf.pages:
                                    text = page.extract_text()
                                    if text:
                                        pdf_text += text

                        except Exception as e:
                            logging.error("PDF read error: " + str(e))

                        # ============================================
                        # CLASSIFICATION
                        # ============================================

                        if "invoice" in pdf_text.lower():
                            classification = "INCOMING"
                            destination_folder = incoming_folder
                        else:
                            classification = "REJECTED"
                            destination_folder = rejected_folder

                        final_path = os.path.join(destination_folder, final_name)

                        os.rename(temp_path, final_path)

                        logging.info(f"{final_name} moved to {classification}")

            # Mark email as read
            mail.store(email_id, '+FLAGS', '\\Seen')

        except Exception as e:
            logging.error("Email processing error: " + str(e))

    mail.logout()

    logging.info("Processing completed successfully")
    logging.info("==========================================")

    return "Processing Completed Successfully"