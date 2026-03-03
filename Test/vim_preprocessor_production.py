"""
SAP VIM Production Email Preprocessor
Author: Production Version
"""

import imaplib
import email
import os
import io
import json
import logging
import hashlib
import shutil
import time
import re
from datetime import datetime
from email.header import decode_header

import pdfplumber

# Optional OCR
OCR_ENABLED = False
try:
    import pytesseract
    from pdf2image import convert_from_bytes
    OCR_ENABLED = True
except:
    pass


# ============================================================
# CONFIGURATION (USE ENV VARIABLES IN PRODUCTION)
# ============================================================

IMAP_SERVER = os.environ.get("VIM_IMAP_SERVER", "imap.one.com")
IMAP_PORT = int(os.environ.get("VIM_IMAP_PORT", "993"))

EMAIL_USER = os.environ.get("VIM_EMAIL_USER", "divyanshu.purohit@dpe-technologies.com")
EMAIL_PASS = os.environ.get("VIM_EMAIL_PASS", "Summer2024*")

WATCH_FOLDER = os.environ.get(
    "VIM_WATCH_FOLDER",
    r"C:\VIM_AUTOMATION\incoming"
)

REJECT_FOLDER = r"C:\VIM_AUTOMATION\rejected"
LOG_FOLDER = r"C:\VIM_AUTOMATION\logs"
STATE_FILE = r"C:\VIM_AUTOMATION\processed_hashes.json"


# Trusted vendor domains
TRUSTED_DOMAINS = [
    "vendor.com",
    "supplier.com",
    "dpe-technologies.com"
]


# ============================================================
# LOGGING SETUP
# ============================================================

os.makedirs(LOG_FOLDER, exist_ok=True)

logging.basicConfig(
    filename=os.path.join(LOG_FOLDER, "vim_preprocessor.log"),
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)

print("Starting SAP VIM production preprocessor")


# ============================================================
# STATE MANAGEMENT (duplicate detection)
# ============================================================

def load_state():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r") as f:
            return json.load(f)
    return {}


def save_state(state):
    with open(STATE_FILE, "w") as f:
        json.dump(state, f, indent=2)


def compute_hash(data):
    return hashlib.sha256(data).hexdigest()


processed_hashes = load_state()


# ============================================================
# MAIL CONNECTION
# ============================================================

def connect_mailbox():

    for attempt in range(3):

        try:
            mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
            mail.login(EMAIL_USER, EMAIL_PASS)
            mail.select("INBOX")

            logging.info("Connected to mailbox")
            return mail

        except Exception as e:
            logging.error(f"Connection failed attempt {attempt}: {e}")
            time.sleep(5)

    raise Exception("Mailbox connection failed")


# ============================================================
# EMAIL PARSING
# ============================================================

def parse_email(mail, email_id):

    _, msg_data = mail.fetch(email_id, "(RFC822)")
    raw = msg_data[0][1]

    msg = email.message_from_bytes(raw)

    subject = decode_header(msg["Subject"])[0][0]
    if isinstance(subject, bytes):
        subject = subject.decode()

    sender = msg.get("From")

    attachments = []

    for part in msg.walk():

        filename = part.get_filename()

        if filename and filename.lower().endswith(".pdf"):

            attachments.append({
                "filename": filename,
                "data": part.get_payload(decode=True)
            })

    return {
        "sender": sender,
        "subject": subject,
        "attachments": attachments
    }


# ============================================================
# VENDOR VALIDATION
# ============================================================

def validate_vendor(sender):

    sender = sender.lower()

    for domain in TRUSTED_DOMAINS:

        if domain in sender:
            return True

    return False


# ============================================================
# PDF TEXT EXTRACTION
# ============================================================

def extract_pdf_text(pdf_bytes):

    text = ""

    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:

            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text += t

    except:
        pass

    if not text and OCR_ENABLED:

        try:
            images = convert_from_bytes(pdf_bytes)
            for img in images:
                text += pytesseract.image_to_string(img)

        except:
            pass

    return text


# ============================================================
# FIELD EXTRACTION
# ============================================================

def extract_invoice_number(text):

    match = re.search(
        r"Invoice\s*(?:No|Number|#)[:\s]*([A-Z0-9\-]+)",
        text,
        re.IGNORECASE
    )

    if match:
        return match.group(1)

    return None


# ============================================================
# SAVE TO VIM WATCH FOLDER
# ============================================================

def save_to_vim(pdf_bytes, filename, metadata):

    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")

    base = f"{timestamp}_{filename}"

    temp_pdf = os.path.join(WATCH_FOLDER, base + ".tmp")
    final_pdf = os.path.join(WATCH_FOLDER, base)

    os.makedirs(WATCH_FOLDER, exist_ok=True)

    with open(temp_pdf, "wb") as f:
        f.write(pdf_bytes)

    os.rename(temp_pdf, final_pdf)

    meta_file = final_pdf + ".json"

    with open(meta_file, "w") as f:
        json.dump(metadata, f, indent=2)

    logging.info(f"Saved to VIM: {final_pdf}")


# ============================================================
# REJECT FILE
# ============================================================

def reject_file(pdf_bytes, filename, reason):

    os.makedirs(REJECT_FOLDER, exist_ok=True)

    path = os.path.join(REJECT_FOLDER, filename)

    with open(path, "wb") as f:
        f.write(pdf_bytes)

    logging.warning(f"Rejected {filename} : {reason}")


# ============================================================
# MAIN PROCESSOR
# ============================================================

def process():

    mail = connect_mailbox()

    _, messages = mail.search(None, "UNSEEN")

    ids = messages[0].split()

    logging.info(f"Found {len(ids)} emails")

    for email_id in ids:

        try:

            email_data = parse_email(mail, email_id)

            sender = email_data["sender"]

            if not validate_vendor(sender):

                logging.warning("Untrusted vendor")
                continue

            for attachment in email_data["attachments"]:

                data = attachment["data"]
                filename = attachment["filename"]

                file_hash = compute_hash(data)

                if file_hash in processed_hashes:
                    logging.warning("Duplicate detected")
                    continue

                text = extract_pdf_text(data)

                invoice_number = extract_invoice_number(text)

                metadata = {
                    "sender": sender,
                    "invoice_number": invoice_number,
                    "processed_time": datetime.now().isoformat()
                }

                save_to_vim(data, filename, metadata)

                processed_hashes[file_hash] = True

                save_state(processed_hashes)

            mail.store(email_id, "+FLAGS", "\\Seen")

        except Exception as e:

            logging.error(f"Processing error: {e}")

    mail.logout()


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":

    process()

    print("Processing complete")