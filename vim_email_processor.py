# ============================================================
# SAP VIM Email Processor
# ============================================================

import imaplib
import email
import email.utils
import os
import re
import pdfplumber
import logging
import pikepdf

from datetime import datetime, timedelta
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors
from io import BytesIO


IMAP_SERVER = "imap.one.com"
IMAP_PORT   = 993

# ============================================================
# INVOICE KEYWORDS
# ============================================================

INVOICE_KEYWORDS = [
    "invoice", "inv #", "inv#", "inv-", "invoice no", "invoice number",
    "rechnung", "factura", "fattura", "faktura", "bill to", "amount due",
    "payment due", "tax invoice", "pro forma", "proforma", "remittance",
]

def _is_invoice(text: str) -> bool:
    t = text.lower()
    return any(kw in t for kw in INVOICE_KEYWORDS)


# ============================================================
# CLEAN FILENAME
# ============================================================

def _clean_filename(name: str, max_len: int = 80) -> str:
    name = re.sub(r'[\\/*?:"<>|]', "_", name)
    name = re.sub(r'\s+', "_", name.strip())
    name = re.sub(r'_+', "_", name)
    base, ext = os.path.splitext(name)
    if len(base) > max_len:
        base = base[:max_len]
    return base + ext


# ============================================================
# THREAD-SAFE FILE LOGGER
# ============================================================

def _setup_logging(log_file: str):
    root = logging.getLogger()
    for h in root.handlers[:]:
        root.removeHandler(h)
        try:
            h.close()
        except Exception:
            pass

    os.makedirs(os.path.dirname(log_file), exist_ok=True)
    fh = logging.FileHandler(log_file, mode="a", encoding="utf-8")
    fh.setFormatter(logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S"
    ))
    sh = logging.StreamHandler()
    sh.setFormatter(logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S"
    ))
    root.setLevel(logging.INFO)
    root.addHandler(fh)
    root.addHandler(sh)


# ============================================================
# PDF/A CONVERSION  (pikepdf)
# ============================================================

def convert_to_pdfa(input_pdf: str, output_pdf: str) -> bool:
    try:
        with pikepdf.open(input_pdf) as pdf:
            xmp = (
                '<?xpacket begin="\xef\xbb\xbf" id="W5M0MpCehiHzreSzNTczkc9d"?>'
                '<x:xmpmeta xmlns:x="adobe:ns:meta/">'
                '<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">'
                '<rdf:Description rdf:about="" '
                'xmlns:pdfaid="http://www.aiim.org/pdfa/ns/id/">'
                '<pdfaid:part>1</pdfaid:part>'
                '<pdfaid:conformance>B</pdfaid:conformance>'
                '</rdf:Description></rdf:RDF></x:xmpmeta>'
                '<?xpacket end="w"?>'
            )
            ms = pikepdf.Stream(pdf, xmp.encode("utf-8"))
            ms["/Type"]    = pikepdf.Name("/Metadata")
            ms["/Subtype"] = pikepdf.Name("/XML")
            pdf.Root["/Metadata"] = ms

            oi = pdf.make_indirect(pikepdf.Dictionary(
                Type=pikepdf.Name("/OutputIntent"),
                S=pikepdf.Name("/GTS_PDFA1"),
                OutputConditionIdentifier=pikepdf.String("sRGB"),
                RegistryName=pikepdf.String("http://www.color.org"),
                Info=pikepdf.String("sRGB IEC61966-2.1"),
            ))
            pdf.Root["/OutputIntents"] = pikepdf.Array([oi])
            pdf.save(output_pdf, linearize=True)

        if input_pdf != output_pdf and os.path.exists(input_pdf):
            os.remove(input_pdf)
        return True

    except Exception as e:
        logging.warning(f"PDF/A conversion failed: {e} — keeping original")
        import shutil
        if input_pdf != output_pdf:
            shutil.copy2(input_pdf, output_pdf)
            try:
                os.remove(input_pdf)
            except Exception:
                pass
        return False


# ============================================================
# EMAIL → PDF  (ReportLab Platypus)
# ============================================================

def build_email_pdf(sender: str, subject: str,
                    date_str: str, body: str,
                    output_path: str):

    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=22*mm, rightMargin=22*mm,
        topMargin=20*mm,  bottomMargin=20*mm,
    )

    def _s(fname, size, bold=False, color="#111111", space_after=4):
        return ParagraphStyle(
            fname,
            fontName="Helvetica-Bold" if bold else "Helvetica",
            fontSize=size,
            textColor=colors.HexColor(color),
            spaceAfter=space_after,
            leading=size * 1.45,
        )

    s_title = _s("t", 13, bold=True, color="#0f2b46", space_after=6)
    s_label = _s("l",  8, bold=True, color="#555577", space_after=2)
    s_value = _s("v",  9,            color="#222222", space_after=8)
    s_body  = _s("b",  9,            color="#111111", space_after=3)

    def safe(t):
        return (t or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    story = [
        Paragraph("SAP VIM — Email Document", s_title),
        HRFlowable(width="100%", thickness=1,
                   color=colors.HexColor("#ccddee"), spaceAfter=8),
        Spacer(1, 3*mm),
        Paragraph("From",    s_label), Paragraph(safe(sender),   s_value),
        Paragraph("Subject", s_label), Paragraph(safe(subject),  s_value),
        Paragraph("Date",    s_label), Paragraph(safe(date_str), s_value),
        HRFlowable(width="100%", thickness=0.5,
                   color=colors.HexColor("#dddddd"), spaceAfter=6),
        Spacer(1, 3*mm),
        Paragraph("Message Body", s_label),
        Spacer(1, 2*mm),
    ]

    for line in body.split("\n"):
        sl = safe(line)
        story.append(Paragraph(sl, s_body) if sl.strip() else Spacer(1, 2*mm))

    doc.build(story)
    with open(output_path, "wb") as f:
        f.write(buf.getvalue())


# ============================================================
# IMAP HELPERS
# ============================================================

def _imap_date(dt: datetime) -> str:
    return dt.strftime("%d-%b-%Y")

def _parse_email_date(date_str: str):
    if not date_str:
        return None
    try:
        return email.utils.parsedate_to_datetime(date_str).replace(tzinfo=None)
    except Exception:
        return None


# ============================================================
# MAIN PROCESSOR
# ============================================================

def run_processor(email_user: str,
                  email_pass: str,
                  incoming_folder: str,
                  rejected_folder: str,
                  log_file: str,
                  start_date: str,
                  end_date: str) -> str:

    _setup_logging(log_file)
    logging.info("=" * 50)
    logging.info("VIM Email Processing Started")
    logging.info(f"Range: {start_date} → {end_date}")
    logging.info("=" * 50)

    try:
        dt_start = datetime.strptime(start_date, "%Y-%m-%d")
        dt_end   = datetime.strptime(end_date,   "%Y-%m-%d")
    except ValueError as e:
        msg = f"Invalid date: {e}"
        logging.error(msg)
        return msg

    if dt_start > dt_end:
        msg = "Start date must not be after end date."
        logging.error(msg)
        return msg

    try:
        logging.info(f"Connecting to {IMAP_SERVER} ...")
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(email_user, email_pass)
        mail.select("INBOX")
        logging.info("Mailbox connected OK")
    except Exception as e:
        logging.error(f"Login failed: {e}")
        return f"Login failed: {e}"

    before_dt  = dt_end + timedelta(days=1)
    search_str = f'(SINCE "{_imap_date(dt_start)}" BEFORE "{_imap_date(before_dt)}")'
    logging.info(f"IMAP search: {search_str}")

    status, messages = mail.search(None, search_str)
    if status != "OK":
        mail.logout()
        return "IMAP search failed."

    email_ids  = messages[0].split()
    logging.info(f"Found {len(email_ids)} email(s) in range")

    processed  = 0
    incoming_n = 0
    rejected_n = 0

    for email_id in email_ids:
        try:
            _, msg_data = mail.fetch(email_id, "(RFC822)")
            msg = email.message_from_bytes(msg_data[0][1])
        except Exception as e:
            logging.error(f"Fetch error id={email_id}: {e}")
            continue

        sender   = msg.get("From",    "Unknown")
        subject  = msg.get("Subject", "No Subject")
        date_hdr = msg.get("Date",    "")

        edt = _parse_email_date(date_hdr)
        if edt and not (dt_start <= edt <= dt_end + timedelta(days=1)):
            logging.info(f"Skip (out of range): {subject[:50]}")
            continue

        logging.info(f"Processing: {subject[:60]}")

        pdf_found = False

        # ── CASE 1: PDF attachment ────────────────────────────────────
        if msg.is_multipart():
            for part in msg.walk():
                fname = part.get_filename()
                if not fname or not fname.lower().endswith(".pdf"):
                    continue

                pdf_found = True
                clean = _clean_filename(fname)
                temp  = os.path.join(incoming_folder, f"tmp_{clean}")

                try:
                    with open(temp, "wb") as f:
                        f.write(part.get_payload(decode=True))
                    logging.info(f"  Attachment: {clean}")

                    text = ""
                    with pdfplumber.open(temp) as p:
                        for pg in p.pages:
                            t = pg.extract_text()
                            if t:
                                text += t

                    if _is_invoice(text + " " + subject):
                        dest = os.path.join(incoming_folder, clean)
                        logging.info("  [INVOICE] → incoming")
                        incoming_n += 1
                    else:
                        dest = os.path.join(rejected_folder, clean)
                        logging.info("  [NOT INVOICE] → rejected")
                        rejected_n += 1

                    convert_to_pdfa(temp, dest)

                except Exception as e:
                    logging.error(f"  PDF error ({clean}): {e}")
                    import shutil
                    try:
                        shutil.move(temp, os.path.join(rejected_folder, clean))
                    except Exception:
                        pass

        # ── CASE 2: No PDF → email body → PDF/A ──────────────────────
        if not pdf_found:
            logging.info("  No attachment — converting email to PDF")

            body_text = ""
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        try:
                            body_text += part.get_payload(
                                decode=True).decode(errors="ignore")
                        except Exception:
                            pass
            else:
                try:
                    body_text = msg.get_payload(
                        decode=True).decode(errors="ignore")
                except Exception:
                    body_text = ""

            ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_sub = _clean_filename(subject or "email")[:50]
            filename = f"email_{ts}_{safe_sub}.pdf"
            temp_pdf = os.path.join(incoming_folder, f"tmp_{filename}")

            try:
                build_email_pdf(
                    sender=sender, subject=subject,
                    date_str=date_hdr, body=body_text,
                    output_path=temp_pdf,
                )

                if _is_invoice(body_text + " " + subject):
                    final = os.path.join(incoming_folder, filename)
                    logging.info(f"  [INVOICE] in body → incoming: {filename}")
                    incoming_n += 1
                else:
                    final = os.path.join(rejected_folder, filename)
                    logging.info(f"  [NOT INVOICE] → rejected: {filename}")
                    rejected_n += 1

                convert_to_pdfa(temp_pdf, final)

            except Exception as e:
                logging.error(f"  Email→PDF failed: {e}")

        processed += 1

    mail.logout()

    summary = (
        f"Done. {processed} emails processed — "
        f"{incoming_n} incoming, {rejected_n} rejected."
    )
    logging.info("=" * 50)
    logging.info(summary)
    logging.info("=" * 50)
    return summary
