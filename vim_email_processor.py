import imaplib, email, os, re, shutil, subprocess, tempfile, html, json, hashlib
from email.header import decode_header, make_header
from email.utils import parseaddr
from datetime import datetime, timedelta

import pdfplumber
from pypdf import PdfWriter
from PIL import Image, ImageOps
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Optional OCR (needs tesseract-ocr + poppler-utils on the server)
OCR_ENABLED = False
try:
    import pytesseract
    from pdf2image import convert_from_path
    OCR_ENABLED = True
except Exception:
    pass

IMAP_SERVER = "imap.one.com"
IMAP_PORT = 993

# ---------- SENDER FILTER ----------
# Only emails from these domains are processed. Edit this list, or set the
# VIM_TRUSTED_DOMAINS env var (comma separated). Empty value = allow everyone.
TRUSTED_DOMAINS = [
    d.strip().lower()
    for d in os.environ.get(
        "VIM_TRUSTED_DOMAINS",
        "vendor.com,supplier.com,dpe-technologies.com,turnpikes.pro"
    ).split(",")
    if d.strip()
]

OCR_MAX_PAGES = 5

OFFICE_EXT = {".doc", ".docx", ".xls", ".xlsx", ".xlsm", ".odt", ".ods", ".rtf"}
IMAGE_EXT = {".jpg", ".jpeg", ".png"}
SUPPORTED = {".pdf"} | OFFICE_EXT | IMAGE_EXT

# Unicode font (optional): installed via fonts-dejavu-core
FONT = "Helvetica"
_dejavu = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
if os.path.exists(_dejavu):
    pdfmetrics.registerFont(TTFont("DejaVu", _dejavu))
    FONT = "DejaVu"


# ---------- helpers ----------

def clean_filename(name):
    name = re.sub(r"[^\w\-_\. ]", "_", name)
    return name.replace(" ", "_")


def decode_str(value):
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return str(value)


INVOICE_PATTERN = re.compile(
    r"\b("
    r"invoices?|inv|tax\s*invoice|proforma|pro-forma|"
    r"credit\s*note|debit\s*note|payment\s*advice|remittance|"
    r"bill\s*(no|number|#)|billing\s*statement|"
    r"amount\s*due|total\s*due|payment\s*due|due\s*date"
    r")\b",
    re.IGNORECASE,
)


def is_invoice(text):
    return bool(INVOICE_PATTERN.search(text))


# ---------- SENDER FILTER ----------

def is_trusted(from_header):
    if not TRUSTED_DOMAINS:
        return True
    addr = parseaddr(from_header or "")[1].lower()
    domain = addr.rsplit("@", 1)[-1] if "@" in addr else ""
    return any(domain == d or domain.endswith("." + d) for d in TRUSTED_DOMAINS)


# ---------- INVOICE NUMBER EXTRACTION ----------

INVOICE_NO_PATTERN = re.compile(
    r"\b(?:invoice|inv)\s*(?:no\.?|number|num|#)\s*[:#.\-]?\s*([A-Z0-9][A-Z0-9\-/_]{1,39})",
    re.IGNORECASE,
)


def extract_invoice_number(text):
    for m in INVOICE_NO_PATTERN.finditer(text):
        token = m.group(1).strip("-/_")
        if any(c.isdigit() for c in token):   # real numbers contain a digit
            return token
    return None


# ---------- DUPLICATE CHECK ----------

def load_state(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_state(path, state):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)
    os.replace(tmp, path)


def make_fingerprint(att_hashes, sender, subject, body):
    if att_hashes:   # same attachment bytes = same invoice
        raw = "|".join(sorted(att_hashes))
    else:            # body-only email
        raw = f"{sender}|{subject}|{body.strip()}"
    return hashlib.sha256(raw.encode("utf-8", "replace")).hexdigest()


# ---------- body extraction ----------

def extract_body(msg):
    plain, html_body = "", ""
    for part in msg.walk():
        if part.is_multipart() or part.get_filename():
            continue
        ctype = part.get_content_type()
        if ctype not in ("text/plain", "text/html"):
            continue
        payload = part.get_payload(decode=True) or b""
        text = payload.decode(part.get_content_charset() or "utf-8", errors="replace")
        if ctype == "text/plain":
            plain += text
        else:
            html_body += text
    if plain.strip():
        return plain
    if html_body.strip():  # HTML-only email: strip tags
        t = re.sub(r"(?is)<(script|style).*?</\1>", "", html_body)
        t = re.sub(r"(?i)<br\s*/?>|</p>|</div>|</tr>", "\n", t)
        t = re.sub(r"<[^>]+>", "", t)
        return html.unescape(t)
    return ""


# ---------- converters ----------

def body_to_pdf(sender, subject, date, body, out_path):
    styles = getSampleStyleSheet()
    for s in styles.byName.values():
        s.fontName = FONT
    doc = SimpleDocTemplate(out_path, pagesize=A4)
    esc = lambda s: html.escape(s).replace("\n", "<br/>")
    story = [
        Paragraph(f"<b>From:</b> {esc(sender)}", styles["Normal"]),
        Paragraph(f"<b>Date:</b> {esc(date)}", styles["Normal"]),
        Paragraph(f"<b>Subject:</b> {esc(subject)}", styles["Normal"]),
        Spacer(1, 16),
        Paragraph(esc(body), styles["Normal"]),
    ]
    doc.build(story)


def office_to_pdf(src, outdir):
    profile = tempfile.mkdtemp(prefix="lo_profile_")
    try:
        subprocess.run(
            ["soffice", f"-env:UserInstallation=file://{profile}",
             "--headless", "--convert-to", "pdf", "--outdir", outdir, src],
            check=True, timeout=180, capture_output=True)
    finally:
        shutil.rmtree(profile, ignore_errors=True)
    return os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + ".pdf")


def image_to_pdf(src, outdir):
    out = os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + ".pdf")
    img = ImageOps.exif_transpose(Image.open(src)).convert("RGB")
    img.save(out, "PDF", resolution=150)
    return out


def to_pdf(src, outdir):
    ext = os.path.splitext(src)[1].lower()
    if ext == ".pdf":
        return src
    if ext in OFFICE_EXT:
        return office_to_pdf(src, outdir)
    if ext in IMAGE_EXT:
        return image_to_pdf(src, outdir)
    raise ValueError(f"Unsupported type {ext}")


def convert_to_pdfa(input_pdf, output_pdf):
    try:
        subprocess.run([
            "gs", "-dPDFA=2", "-dBATCH", "-dNOPAUSE", "-dQUIET",
            "-sDEVICE=pdfwrite",
            "-dPDFACompatibilityPolicy=1",
            "-sColorConversionStrategy=UseDeviceIndependentColor",
            f"-sOutputFile={output_pdf}", input_pdf
        ], check=True, timeout=180)
    except Exception as e:
        print("PDF/A conversion failed, saving original PDF:", e)
        shutil.copyfile(input_pdf, output_pdf)


# ---------- text extraction (with OCR fallback) ----------

def ocr_pdf(path):
    if not OCR_ENABLED:
        return ""
    try:
        pages = convert_from_path(path, dpi=200, first_page=1, last_page=OCR_MAX_PAGES)
        return "\n".join(pytesseract.image_to_string(p) for p in pages)
    except Exception as e:
        print("OCR failed:", e)
        return ""


def pdf_text(path):
    text = ""
    try:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                text += (page.extract_text() or "") + "\n"
    except Exception:
        pass
    if len(text.strip()) < 20:      # scanned PDF / image -> OCR
        text += "\n" + ocr_pdf(path)
    return text


# ---------- per-email processing ----------
# returns: "saved", "duplicate", "untrusted" or "not_invoice"

def process_message(msg, uid, incoming, state):
    # 1. sender filter
    if not is_trusted(msg.get("From", "")):
        return "untrusted"

    subject = decode_str(msg.get("Subject"))
    sender = decode_str(msg.get("From"))
    date = msg.get("Date", "")
    body = extract_body(msg)

    # 2. collect attachments
    atts = []   # (safe_name, ext, data)
    for part in msg.walk():
        if part.is_multipart():
            continue
        filename = decode_str(part.get_filename())
        if not filename:
            continue
        ext = os.path.splitext(filename)[1].lower()
        if ext not in SUPPORTED:
            continue
        data = part.get_payload(decode=True)
        if not data:
            continue
        # skip signature logos / tracking pixels
        if ext in IMAGE_EXT and (
            "inline" in str(part.get("Content-Disposition", "")).lower()
            or len(data) < 15 * 1024
        ):
            continue
        atts.append((clean_filename(filename), ext, data))

    # 3. duplicate check
    att_hashes = [hashlib.sha256(d).hexdigest() for _, _, d in atts]
    fp = make_fingerprint(att_hashes, sender, subject, body)
    prev = state.get(fp)
    if prev and prev.get("file") and os.path.exists(os.path.join(incoming, prev["file"])):
        return "duplicate"

    with tempfile.TemporaryDirectory() as tmp:
        outputs = []      # list of (label, pdf_path)
        all_text = subject + "\n" + body

        for n, (safe, ext, data) in enumerate(atts, 1):
            src_dir = os.path.join(tmp, f"att{n}")
            os.makedirs(src_dir)
            src = os.path.join(src_dir, safe)
            with open(src, "wb") as f:
                f.write(data)
            try:
                pdf = to_pdf(src, src_dir)
            except Exception as e:
                print(f"Skipping attachment {safe}: {e}")
                continue
            outputs.append((safe, pdf))
            all_text += "\n" + pdf_text(pdf)

        # 4. skip emails that are not invoice related
        if not is_invoice(all_text):
            return "not_invoice"

        # body PDF (skipped if the mail has no body text)
        if body.strip():
            body_pdf = os.path.join(tmp, "body.pdf")
            body_to_pdf(sender, subject, date, body, body_pdf)
            outputs.insert(0, ("body", body_pdf))

        if not outputs:
            return "not_invoice"

        # merge: body first, then attachments, into ONE pdf
        writer = PdfWriter()
        added = 0
        for _, pdf in outputs:
            try:
                writer.append(pdf)
                added += 1
            except Exception as e:
                print(f"Could not merge {pdf}: {e}")
        if not added:
            return "not_invoice"

        merged = os.path.join(tmp, "merged.pdf")
        with open(merged, "wb") as f:
            writer.write(f)

        # 5. invoice number -> used in file name
        inv_no = extract_invoice_number(all_text)
        if inv_no:
            base = f"{inv_no}_{uid}"
        else:
            base = f"{uid}_{clean_filename(subject)[:50] or 'email'}"
        name = clean_filename(base)[:120] + ".pdf"

        convert_to_pdfa(merged, os.path.join(incoming, name))

        state[fp] = {
            "file": name,
            "invoice_number": inv_no,
            "sender": sender,
            "subject": subject,
            "processed_time": datetime.now().isoformat(),
        }
        return "saved"


# ---------- main ----------

def run_processor(email_user, email_pass, incoming_folder, rejected_folder,
                  start_date, end_date, mail_type):
    start = datetime.strptime(start_date, "%Y-%m-%d")
    end = datetime.strptime(end_date, "%Y-%m-%d") + timedelta(days=1)
    s, e = start.strftime("%d-%b-%Y"), end.strftime("%d-%b-%Y")

    flag = {"unread": "UNSEEN ", "read": "SEEN "}.get(mail_type, "")
    query = f'({flag}SINCE "{s}" BEFORE "{e}")'

    # duplicate/metadata file lives in the user's folder, NOT in incoming
    # (so it does not show up in the dashboard counts)
    state_path = os.path.join(os.path.dirname(incoming_folder), "processed_hashes.json")
    state = load_state(state_path)

    counts = {"saved": 0, "duplicate": 0, "untrusted": 0, "not_invoice": 0, "failed": 0}

    mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    try:
        mail.login(email_user, email_pass)
        mail.select("INBOX")
        _, data = mail.uid("search", None, query)

        for uid in data[0].split():
            try:
                _, msg_data = mail.uid("fetch", uid, "(BODY.PEEK[])")
                msg = email.message_from_bytes(msg_data[0][1])
                status = process_message(msg, uid.decode(), incoming_folder, state)
                if status == "saved":
                    save_state(state_path, state)
                    mail.uid("store", uid, "+FLAGS", "\\Seen")
                counts[status] += 1
            except Exception as ex:
                counts["failed"] += 1
                print(f"Email {uid} failed: {ex}")
    finally:
        try:
            mail.logout()
        except Exception:
            pass

    return (f"{counts['saved']} invoice emails saved, "
            f"{counts['duplicate']} duplicates skipped, "
            f"{counts['untrusted']} untrusted sender skipped, "
            f"{counts['not_invoice']} not invoice skipped, "
            f"{counts['failed']} failed.")
