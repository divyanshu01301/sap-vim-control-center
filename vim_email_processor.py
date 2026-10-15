import imaplib, email, os, re, shutil, subprocess, tempfile, html
from email.header import decode_header, make_header
from datetime import datetime, timedelta

import pdfplumber
from PIL import Image, ImageOps
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

IMAP_SERVER = "imap.one.com"
IMAP_PORT = 993

OFFICE_EXT = {".doc", ".docx", ".xls", ".xlsx", ".xlsm", ".odt", ".ods", ".rtf"}
IMAGE_EXT = {".jpg", ".jpeg", ".png"}
SUPPORTED = {".pdf"} | OFFICE_EXT | IMAGE_EXT

# Unicode font (optional): apt install fonts-dejavu-core
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


def is_invoice(text):
    return bool(re.search(r"\b(invoice|inv)\s*(no|number|#|num)?\b", text.lower()))


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


def pdf_text(path):
    text = ""
    try:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                text += page.extract_text() or ""
    except Exception:
        pass
    return text


# ---------- per-email processing ----------

def process_message(msg, uid, incoming, rejected):
    subject = decode_str(msg.get("Subject"))
    sender = decode_str(msg.get("From"))
    date = msg.get("Date", "")
    body = extract_body(msg)

    with tempfile.TemporaryDirectory() as tmp:
        outputs = []      # (final_name, pdf_path)
        all_text = subject + "\n" + body

        # attachments
        n = 0
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

            n += 1
            safe = clean_filename(filename)
            src_dir = os.path.join(tmp, f"att{n}")
            os.makedirs(src_dir)
            src = os.path.join(src_dir, safe)
            with open(src, "wb") as f:
                f.write(data)
            try:
                pdf = to_pdf(src, src_dir)
            except Exception as e:
                print(f"Skipping attachment {filename}: {e}")
                continue
            stem = os.path.splitext(safe)[0]
            outputs.append((f"{uid}_{n}_{stem}.pdf", pdf))
            all_text += "\n" + pdf_text(pdf)

        # body PDF: only if there is a body (case 1 and 2); skipped for attachment-only mails
        if body.strip():
            body_pdf = os.path.join(tmp, "body.pdf")
            body_to_pdf(sender, subject, date, body, body_pdf)
            outputs.insert(0, (f"{uid}_body.pdf", body_pdf))

        if not outputs:
            return 0

        target = incoming if is_invoice(all_text) else rejected
        for name, pdf in outputs:
            convert_to_pdfa(pdf, os.path.join(target, name))
        return len(outputs)


# ---------- main ----------

def run_processor(email_user, email_pass, incoming_folder, rejected_folder,
                  start_date, end_date, mail_type):
    start = datetime.strptime(start_date, "%Y-%m-%d")
    end = datetime.strptime(end_date, "%Y-%m-%d") + timedelta(days=1)
    s, e = start.strftime("%d-%b-%Y"), end.strftime("%d-%b-%Y")

    flag = {"unread": "UNSEEN ", "read": "SEEN "}.get(mail_type, "")
    query = f'({flag}SINCE "{s}" BEFORE "{e}")'

    emails_done, files_made, failed = 0, 0, 0
    mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    try:
        mail.login(email_user, email_pass)
        mail.select("INBOX")
        _, data = mail.uid("search", None, query)

        for uid in data[0].split():
            try:
                _, msg_data = mail.uid("fetch", uid, "(BODY.PEEK[])")
                msg = email.message_from_bytes(msg_data[0][1])
                files_made += process_message(msg, uid.decode(), incoming_folder, rejected_folder)
                mail.uid("store", uid, "+FLAGS", "\\Seen")
                emails_done += 1
            except Exception as ex:
                failed += 1
                print(f"Email {uid} failed: {ex}")
    finally:
        try:
            mail.logout()
        except Exception:
            pass

    return f"{emails_done} emails processed, {files_made} PDF/A files created, {failed} failed."
