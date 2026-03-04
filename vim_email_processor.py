# ============================================================
# SAP VIM Email Processor
# Email / PDF → PDF/A + Invoice Classification
# ============================================================

import imaplib
import email
import os
import pdfplumber
import logging
import subprocess
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
# PDF → PDF/A
# ============================================================

def convert_to_pdfa(input_pdf, output_pdf):

    try:

        subprocess.run([
            "gs",
            "-dPDFA",
            "-dBATCH",
            "-dNOPAUSE",
            "-sDEVICE=pdfwrite",
            f"-sOutputFile={output_pdf}",
            input_pdf
        ], check=True)

        os.remove(input_pdf)

    except Exception as e:

        logging.warning("PDF/A conversion failed, keeping original PDF")


# ============================================================
# MAIN PROCESSOR
# ============================================================

def run_processor(email_user,
                  email_pass,
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

        logging.info("Mailbox connected")

    except Exception as e:

        logging.error(f"Login failed: {str(e)}")
        return f"Login failed: {str(e)}"


    status, messages = mail.search(None, "ALL")
    email_ids = messages[0].split()

    start_date = datetime.strptime(start_date, "%Y-%m-%d")
    end_date = datetime.strptime(end_date, "%Y-%m-%d")

    logging.info(f"{len(email_ids)} emails found")


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

        logging.info(f"Processing email from {sender}")


        body_text = ""
        pdf_found = False


        # ====================================================
        # CASE 1 — PDF ATTACHMENT EXISTS
        # ====================================================

        if msg.is_multipart():

            for part in msg.walk():

                filename = part.get_filename()

                if filename and filename.lower().endswith(".pdf"):

                    pdf_found = True

                    clean_name = filename.replace(" ", "_")

                    temp_path = os.path.join(incoming_folder, "temp_" + clean_name)

                    with open(temp_path, "wb") as f:
                        f.write(part.get_payload(decode=True))

                    logging.info(f"PDF attachment detected: {clean_name}")

                    try:

                        text = ""

                        with pdfplumber.open(temp_path) as pdf:

                            for page in pdf.pages:
                                page_text = page.extract_text()
                                if page_text:
                                    text += page_text


                        # INVOICE CHECK
                        if "invoice" in text.lower() or "inv" in text.lower():

                            final_path = os.path.join(incoming_folder, clean_name)
                            logging.info("Invoice detected")

                        else:

                            final_path = os.path.join(rejected_folder, clean_name)
                            logging.info("PDF rejected (no invoice keyword)")


                        convert_to_pdfa(temp_path, final_path)

                    except Exception as e:

                        logging.error("PDF read error")

                        os.rename(
                            temp_path,
                            os.path.join(rejected_folder, clean_name)
                        )


        # ====================================================
        # CASE 2 — NO PDF ATTACHMENT
        # ====================================================

        if not pdf_found:

            logging.info("No attachment found, converting email to PDF")

            if msg.is_multipart():

                for part in msg.walk():

                    if part.get_content_type() == "text/plain":

                        body_text += part.get_payload(
                            decode=True).decode(errors="ignore")

            else:

                body_text = msg.get_payload(
                    decode=True).decode(errors="ignore")


            full_text = f"""
From: {sender}

Subject: {subject}

Date: {email_date}

Body:

{body_text}
"""


            filename = f"email_{processed_count}.pdf"

            temp_pdf = os.path.join(incoming_folder, "temp_" + filename)

            convert_email_to_pdf(full_text, temp_pdf)


            if "invoice" in full_text.lower() or "inv" in full_text.lower():

                final_pdf = os.path.join(incoming_folder, filename)
                logging.info("Invoice detected in email body")

            else:

                final_pdf = os.path.join(rejected_folder, filename)
                logging.info("Email rejected")


            convert_to_pdfa(temp_pdf, final_pdf)


        processed_count += 1


    mail.logout()

    logging.info("===== Processing Completed =====")

    return f"Processing completed. {processed_count} emails processed."