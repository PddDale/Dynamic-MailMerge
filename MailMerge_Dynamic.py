# -*- coding: utf-8 -*-
"""
Interactive assistant for sending personalized mail merge e-mails through
classic Outlook.

The script detects the accounts and shared mailboxes available in the
user's Outlook profile, asks which one should be used as the sender, and,
if it is a shared mailbox, offers the option to send "on behalf of" it
(via SentOnBehalfOfName + resolution against the Exchange address book).
It also asks for a Word document with the signature template (Word
formatting is preserved when converting to HTML) and, optionally, a logo
image to include in the signature.

Usage: python MailMerge_Dynamic.py
"""

import glob
import importlib
import os
import re
import subprocess
import sys
import time
from collections import defaultdict


def clean_path(text):
    """Accepts the path pasted 'raw', wrapped in quotes ("..."/'...') or
    parentheses ((...)) — common formats when copying paths from Explorer
    or other sources — and returns just the clean path."""
    text = text.strip()
    pairs = [('"', '"'), ("'", "'"), ("(", ")")]
    for open_char, close_char in pairs:
        if len(text) >= 2 and text.startswith(open_char) and text.endswith(close_char):
            text = text[1:-1].strip()
    return text


# ==================== STEP 1: ENVIRONMENT CHECK ====================

def check_dependencies():
    packages = {
        "win32com.client": "pywin32",
        "pandas": "pandas",
        "openpyxl": "openpyxl",
        "docx": "python-docx",
        "PIL": "Pillow",
    }
    missing = []
    for module, pip_package in packages.items():
        try:
            importlib.import_module(module)
        except ImportError:
            missing.append(pip_package)

    if missing:
        print(f"[Environment] Installing missing packages: {', '.join(missing)}...")
        subprocess.run([sys.executable, "-m", "pip", "install", *missing], check=True)
        print(f"[Environment] Installed: {', '.join(missing)}")
    else:
        print("[Environment] All dependencies are already installed.")


def _list_running_processes():
    """Runs 'tasklist' once (with no filter) so the two process checks
    (classic Outlook and New Outlook) don't each spawn their own tasklist
    process."""
    try:
        output = subprocess.run(
            ["tasklist"],
            capture_output=True, text=True, check=False,
        )
        return output.stdout.upper()
    except Exception:
        return ""


def classic_outlook_running(process_list=None):
    if process_list is None:
        process_list = _list_running_processes()
    # Anchored at line start (image name is tasklist's 1st column) so a
    # third-party process whose name merely contains "OUTLOOK.EXE" as a
    # substring (e.g. "SOMEOUTLOOK.EXE") isn't mistaken for the real thing.
    return re.search(r"(?m)^OUTLOOK\.EXE\s", process_list) is not None


def locate_outlook_exe():
    candidates = []
    for base_env_var in ("PROGRAMFILES", "PROGRAMFILES(X86)"):
        base = os.environ.get(base_env_var)
        if not base:
            continue
        candidates.extend(glob.glob(os.path.join(base, "Microsoft Office", "root", "Office*", "OUTLOOK.EXE")))
        candidates.extend(glob.glob(os.path.join(base, "Microsoft Office", "Office*", "OUTLOOK.EXE")))
    for path in candidates:
        if os.path.exists(path):
            return path
    return None


def detect_new_outlook(process_list=None):
    """Heuristic: the "New Outlook" runs as a packaged app (olk.exe /
    WindowsApps), not as classic OUTLOOK.EXE. If only the packaged app
    process exists and no OUTLOOK.EXE is found, warn the user."""
    if process_list is None:
        process_list = _list_running_processes()
    return re.search(r"(?m)^OLK\.EXE\s", process_list) is not None


def check_classic_outlook():
    process_list = _list_running_processes()
    if detect_new_outlook(process_list) and not classic_outlook_running(process_list):
        print("\n[WARNING] The 'New Outlook' was detected running, which does NOT support")
        print("COM/MAPI automation. Please close the New Outlook, open CLASSIC Outlook")
        print("and press Enter to continue.")
        input()
        return

    if classic_outlook_running(process_list):
        print("[Environment] Classic Outlook is already running.")
        return

    print("[Environment] Classic Outlook not detected. Trying to open it automatically...")
    exe = locate_outlook_exe()
    if exe:
        try:
            subprocess.Popen([exe])
            print(f"[Environment] Outlook started from: {exe}")
            print("[Environment] Waiting a few seconds for COM automation to become available...")
            time.sleep(8)
            return
        except Exception as e:
            print(f"[Environment] Failed to start Outlook automatically: {e}")

    print("\n[ACTION REQUIRED] Could not locate/open classic Outlook automatically.")
    print("Please open classic Outlook manually and press Enter to continue.")
    input()


def run_environment_check():
    print("=" * 70)
    print("STEP 1/6 - Environment check")
    print("=" * 70)
    check_dependencies()
    check_classic_outlook()
    print()


# ==================== STEP 2: USER E-MAIL (INFORMATIONAL) ==================

def ask_user_email(available_mailboxes):
    print("=" * 70)
    print("STEP 2/6 - User identification (for the log only)")
    print("=" * 70)
    email_regex = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

    accounts = [m["label"] for m in available_mailboxes if m["type"] == "account"]
    if accounts:
        print("Configured e-mails:")
        for i, email in enumerate(accounts, start=1):
            print(f"  {i}. {email}")
        while True:
            choice = input("Which one should be used (list number or type the e-mail): ").strip()
            if choice.isdigit() and 1 <= int(choice) <= len(accounts):
                email = accounts[int(choice) - 1]
                print(f"[Info] User e-mail recorded for the log: {email}")
                print()
                return email
            if email_regex.match(choice):
                print(f"[Info] User e-mail recorded for the log: {choice}")
                print()
                return choice
            print("Invalid choice. Type the number of a list option or a valid e-mail.")

    while True:
        email = input("What is your main e-mail (e.g.: john.doe@company.com)? ").strip()
        if email_regex.match(email):
            print(f"[Info] User e-mail recorded for the log: {email}")
            print()
            return email
        print("Invalid e-mail. Try again.")


# ==================== STEP 3: SPREADSHEET SELECTION =========================

def ask_spreadsheet_path():
    import pandas as pd

    while True:
        path = clean_path(input("Path to the Excel file (.xlsx/.xls) with the contacts: "))
        if not os.path.exists(path):
            print("File not found. Try again.")
            continue
        if not path.lower().endswith((".xlsx", ".xls")):
            print("The file must be .xlsx or .xls. Try again.")
            continue
        try:
            xls = pd.ExcelFile(path)
            return path, xls
        except Exception as e:
            print(f"Could not read the file ({e}). Try again.")


def choose_sheet(xls):
    sheets = xls.sheet_names
    while True:
        print("\nAvailable sheets:")
        for i, name in enumerate(sheets, start=1):
            print(f"  {i}. {name}")
        choice = input("Choose the sheet (number or name): ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(sheets):
            return sheets[int(choice) - 1]
        if choice in sheets:
            return choice
        print("Invalid choice. Try again.")


def list_filled_columns(df):
    valid_columns = [c for c in df.columns if str(c).strip() and not str(c).startswith("Unnamed:")]
    print("\nAvailable columns:")
    for i, col in enumerate(valid_columns, start=1):
        print(f"  {i}. {col}")
    return valid_columns


def choose_column(df, label):
    while True:
        valid_columns = list_filled_columns(df)
        choice = input(f"Which column contains the {label}? (number or name): ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(valid_columns):
            return valid_columns[int(choice) - 1]
        if choice in valid_columns:
            return choice
        print("Invalid choice. Try again.")


def validate_column_data(df, column, label):
    series = df[column]
    empty_mask = series.isna() | (series.astype(str).str.strip() == "")
    empty_count = int(empty_mask.sum())
    if empty_count == len(series):
        print(f"[WARNING] The column '{column}' chosen for {label} is completely empty.")
    elif empty_count > 0:
        print(f"[WARNING] The column '{column}' ({label}) has {empty_count} empty row(s).")

    multi_mask = series.dropna().astype(str).str.contains("/", regex=False)
    multi_count = int(multi_mask.sum())
    if multi_count > 0:
        print(f"[WARNING] The column '{column}' ({label}) has {multi_count} cell(s) with multiple values")
        print("        separated by '/'. Those rows need to be split into separate rows")
        print("        before sending, to keep the individual personalization.")
        print("        They will be skipped in the automatic sending if not fixed in the spreadsheet.")


def select_spreadsheet():
    print("=" * 70)
    print("STEP 4/6 - Contact spreadsheet selection")
    print("=" * 70)

    path, xls = ask_spreadsheet_path()
    sheet = choose_sheet(xls)
    df = xls.parse(sheet_name=sheet)

    print(f"\nSheet '{sheet}' selected. Now choose the name and e-mail columns.")
    name_column = choose_column(df, "NAMES")
    email_column = choose_column(df, "E-MAILS")

    validate_column_data(df, name_column, "names")
    validate_column_data(df, email_column, "e-mails")
    print()

    return {
        "path": path,
        "sheet": sheet,
        "df": df,
        "name_column": name_column,
        "email_column": email_column,
    }


# ==================== STEP 4: SENDING MAILBOX SELECTION =====================
# The script detects the full accounts and the shared mailboxes/additional
# folders available in the Outlook profile and asks the user which one
# should be used as the sender. If the choice is a shared mailbox (not a
# full account owned by the user), it asks whether sending should be done
# "on behalf of" it (SentOnBehalfOfName), resolving the address against the
# Exchange address book (GAL). If the user prefers not to use that
# permission, sending proceeds through the default Outlook account.

def list_available_mailboxes(namespace):
    mailboxes = []
    try:
        for acc in namespace.Accounts:
            try:
                email = acc.SmtpAddress
            except Exception:
                continue
            if email:
                mailboxes.append({"label": email, "type": "account", "account_obj": acc})
    except Exception:
        pass

    known_labels = {m["label"].lower() for m in mailboxes}
    try:
        for folder in namespace.Folders:
            try:
                name = (folder.Name or "").strip()
            except Exception:
                continue
            if name and name.lower() not in known_labels:
                mailboxes.append({"label": name, "type": "additional_folder", "account_obj": None})
                known_labels.add(name.lower())
    except Exception:
        pass

    return mailboxes


def mailbox_configured_in_profile(available_mailboxes, target):
    """Checks whether 'target' (e-mail or display name) is registered as a
    full account or as an additional folder (shared mailbox/Full Access) in
    the current Outlook profile, reusing the already-fetched mailbox list
    instead of walking the Outlook COM tree again."""
    target_lower = target.strip().lower()

    for mailbox in available_mailboxes:
        label_lower = mailbox["label"].strip().lower()
        if mailbox["type"] == "account":
            if label_lower == target_lower:
                return True
        elif target_lower in label_lower:
            return True

    return False


def resolve_sending_address(namespace, initial_name_or_email):
    def resolve(value):
        recipient = namespace.CreateRecipient(value)
        recipient.Resolve()
        if not recipient.Resolved:
            return None
        type_ = recipient.AddressEntry.Type
        print(f"  Attempt '{value}' -> Resolved | Type: {type_} | Address: {recipient.Address}")
        return recipient, type_

    print("Resolving the sending address against the address book...")
    result = resolve(initial_name_or_email)

    if result is None or result[1] != "EX":
        fallback_name = input(
            f"Could not resolve '{initial_name_or_email}' as a native Exchange address.\n"
            "Type the exact display name as it appears in the address book search "
            "(or just press Enter to try again with the same value): "
        ).strip()
        if fallback_name:
            result = resolve(fallback_name)

    if result is None:
        raise Exception(
            f"Could not resolve '{initial_name_or_email}' in the address book. "
            "Confirm the exact name as it appears in the address book (GAL) search."
        )

    recipient, type_ = result
    if type_ != "EX":
        print(f"\nWARNING: the resolved address is of type '{type_}', not 'EX' (native Exchange).")
        print("This is likely to keep causing a permission error.\n")

    return recipient.Address


def ask_sending_mailbox(namespace, mailboxes):
    print("=" * 70)
    print("STEP 3/6 - Sending mailbox selection")
    print("=" * 70)

    print("Mailboxes detected in the Outlook profile:")
    for i, mailbox in enumerate(mailboxes, start=1):
        type_label = "full account" if mailbox["type"] == "account" else "additional folder / possible shared mailbox"
        print(f"  {i}. {mailbox['label']} ({type_label})")

    while True:
        choice = input("Which mailbox should be used as the sender (list number or type the e-mail/name): ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(mailboxes):
            chosen_mailbox = mailboxes[int(choice) - 1]
            break
        if choice:
            chosen_mailbox = {"label": choice, "type": "manual", "account_obj": None}
            break
        print("Empty value. Type the number of a list option or a valid e-mail/name.")

    if chosen_mailbox["type"] == "account":
        print(f"\n'{chosen_mailbox['label']}' is a full account configured in the profile.")
        print("Sending will use that account directly, no need for 'Send on behalf of'.\n")
        return {"mode": "full_account", "sending_address": chosen_mailbox["label"], "account_obj": chosen_mailbox["account_obj"]}

    print(f"\nChecking whether '{chosen_mailbox['label']}' is configured in the profile...")
    # Re-query Outlook live (instead of only the list preloaded at the start
    # of the assistant) so a mailbox the user added to the profile during
    # this same session isn't misreported as "not detected".
    current_mailboxes = list_available_mailboxes(namespace)
    if mailbox_configured_in_profile(current_mailboxes, chosen_mailbox["label"]):
        print("Shared mailbox/additional folder detected in the profile.")
    else:
        print("It was not detected as an additional folder/account in the profile (it may still work via")
        print("'Send on behalf of' permission, as long as it was granted in Exchange).")

    answer = input(
        f"Do you want to send the e-mails on behalf of '{chosen_mailbox['label']}' (via 'Send on behalf of')? (y/n): "
    ).strip().lower()
    if answer != "y":
        print("OK, sending will use the default Outlook account (without 'Send on behalf of').\n")
        return {"mode": "full_account", "sending_address": None, "account_obj": None}

    resolved_address = resolve_sending_address(namespace, chosen_mailbox["label"])
    print(f"\nFinal address resolved for sending: {resolved_address}\n")
    return {"mode": "shared_mailbox", "sending_address": resolved_address, "account_obj": None}


# ==================== STEP 5: SIGNATURE (DOCX) AND LOGO ======================
# The signature template is provided by the user as a Word document. The
# conversion to HTML is done directly from python-docx runs (it does not use
# a simplification library such as mammoth), to faithfully preserve bold,
# italic, underline, exact text color and images embedded in the document
# itself (e.g., a logo already pasted into the signature).

_ALIGNMENT_CSS = {
    1: "center",   # WD_ALIGN_PARAGRAPH.CENTER
    2: "right",    # WD_ALIGN_PARAGRAPH.RIGHT
    3: "justify",  # WD_ALIGN_PARAGRAPH.JUSTIFY
}


_EMU_PER_PIXEL = 9525  # EMU per pixel at 96 DPI (Word/Office default)


def _extract_run_images(run):
    import base64
    from docx.oxml.ns import qn

    images_html = []
    for drawing in run._element.findall(".//" + qn("w:drawing")):
        width_px = height_px = None
        extent = drawing.find(".//" + qn("wp:extent"))
        if extent is not None:
            cx = extent.get("cx")
            cy = extent.get("cy")
            if cx and cy:
                try:
                    width_px = round(int(cx) / _EMU_PER_PIXEL)
                    height_px = round(int(cy) / _EMU_PER_PIXEL)
                except Exception:
                    width_px = height_px = None

        for blip in drawing.findall(".//" + qn("a:blip")):
            r_id = blip.get(qn("r:embed"))
            if not r_id:
                continue
            try:
                image_part = run.part.related_parts[r_id]
                b64 = base64.b64encode(image_part.blob).decode("ascii")
                dimensions = f' width="{width_px}" height="{height_px}"' if width_px and height_px else ""
                images_html.append(f'<img src="data:{image_part.content_type};base64,{b64}"{dimensions}>')
            except Exception:
                continue
    return images_html


def _run_to_html(run):
    from html import escape

    parts = []
    text = run.text or ""
    if text:
        html_text = escape(text).replace("\n", "<br>")
        styles = []
        color = None
        try:
            if run.font.color is not None and run.font.color.type is not None and run.font.color.rgb is not None:
                color = str(run.font.color.rgb)
        except Exception:
            pass
        if color:
            styles.append(f"color:#{color}")
        try:
            if run.font.size is not None:
                styles.append(f"font-size:{run.font.size.pt}pt")
        except Exception:
            pass
        if run.bold:
            html_text = f"<strong>{html_text}</strong>"
        if run.italic:
            html_text = f"<em>{html_text}</em>"
        if run.underline:
            html_text = f"<u>{html_text}</u>"
        if styles:
            html_text = f'<span style="{";".join(styles)}">{html_text}</span>'
        parts.append(html_text)
    parts.extend(_extract_run_images(run))
    return "".join(parts)


def _iter_paragraph_runs(paragraph):
    """Walks the paragraph's child elements in document order and yields
    (run, url) for every <w:r>. python-docx's `paragraph.runs` only sees
    <w:r> elements that are direct children of <w:p> — runs nested inside a
    <w:hyperlink> (exactly what Word produces when you insert a hyperlink,
    including an e-mail address formatted as a link) are invisible to it and
    disappear from the generated HTML. Here we also descend into
    <w:hyperlink> and resolve the relationship's target URL so neither the
    text nor the link gets lost."""
    from docx.oxml.ns import qn
    from docx.text.run import Run

    for child in paragraph._p.iterchildren():
        if child.tag == qn("w:r"):
            yield Run(child, paragraph), None
        elif child.tag == qn("w:hyperlink"):
            url = None
            r_id = child.get(qn("r:id"))
            if r_id:
                try:
                    url = paragraph.part.rels[r_id].target_ref
                except Exception:
                    url = None
            for grandchild in child.iterchildren():
                if grandchild.tag == qn("w:r"):
                    yield Run(grandchild, paragraph), url


def _extract_docx_paragraphs_html(path):
    """Converts each paragraph of a .docx into HTML while preserving bold,
    italic, underline, color and font size of every text run, plus any
    embedded images (with the document's original width/height) and
    hyperlinks (including e-mails formatted as links)."""
    from html import escape

    from docx import Document

    doc = Document(path)
    items = []  # each item: (is_blank_line, paragraph_html)
    for paragraph in doc.paragraphs:
        parts = []
        for run, url in _iter_paragraph_runs(paragraph):
            html_run = _run_to_html(run)
            if url and html_run:
                html_run = f'<a href="{escape(url)}">{html_run}</a>'
            parts.append(html_run)
        content = "".join(parts)
        alignment = _ALIGNMENT_CSS.get(paragraph.alignment)
        style = f' style="text-align:{alignment};margin:0"' if alignment else ' style="margin:0"'
        if not content.strip():
            # An empty Word paragraph is usually an intentional blank line
            # between signature blocks (e.g. between the name/e-mail and the
            # address). Preserve it as a blank-line marker instead of
            # dropping it — otherwise that visual separation disappears.
            items.append((True, f"<p{style}>&nbsp;</p>"))
        else:
            items.append((False, f"<p{style}>{content}</p>"))

    # Zero out the default margin browsers/e-mail clients apply to every
    # <p> (which is much larger than Word's actual spacing) and rely on the
    # document's own blank-line paragraphs to reproduce the gaps between
    # blocks — so the final e-mail keeps the same spacing as the original
    # docx instead of a "doubled" spacing.
    while items and items[0][0]:
        items.pop(0)
    while items and items[-1][0]:
        items.pop()

    return [html for _, html in items]


def extract_signature_docx(path):
    return "\n".join(_extract_docx_paragraphs_html(path))


def ask_signature_path():
    print("=" * 70)
    print("STEP 5/6 - Signature template (Word) and logo")
    print("=" * 70)
    while True:
        path = clean_path(input("Path to the Word document (.docx) with the signature template: "))
        if not os.path.exists(path) or not path.lower().endswith(".docx"):
            print("File not found or invalid format (use .docx). Try again.")
            continue
        if os.path.getsize(path) == 0:
            print("The file is empty (0 bytes). Open it, add content and save again. Try another path.")
            continue
        try:
            signature_html = extract_signature_docx(path)
        except Exception as e:
            print(f"Could not read the document ({e}). Try again.")
            continue
        if not signature_html.strip():
            print("The signature document is empty. Try again.")
            continue
        return path, signature_html


def signature_already_has_image(signature_html):
    return "<img" in signature_html


def ask_logo(signature_html):
    if signature_already_has_image(signature_html):
        print("[Info] The signature document already contains an embedded image (extracted with its")
        print("original Word size). The external logo step was skipped to avoid duplication.")
        print()
        return None

    answer = input("Is there a company logo image to include in the signature? (y/n): ").strip().lower()
    if answer != "y":
        print()
        return None
    while True:
        path = clean_path(input("Path to the logo image file (PNG/JPG): "))
        if os.path.exists(path) and path.lower().endswith((".png", ".jpg", ".jpeg")):
            print()
            return path
        print("File not found or invalid format (use .png/.jpg/.jpeg). Try again.")


MAX_LOGO_WIDTH_PX = 180
MAX_LOGO_HEIGHT_PX = 60


def calculate_logo_dimensions(path, max_width=MAX_LOGO_WIDTH_PX, max_height=MAX_LOGO_HEIGHT_PX):
    """Reads the image's real size and returns proportionally reduced
    dimensions to fit within the limit (never enlarges the image, only
    prevents a high-resolution logo from looking huge in the signature)."""
    from PIL import Image

    try:
        with Image.open(path) as img:
            width, height = img.size
    except Exception:
        return None
    if width <= 0 or height <= 0:
        return None

    factor = min(1.0, max_width / width, max_height / height)
    return max(1, round(width * factor)), max(1, round(height * factor))


def build_final_signature(signature_html, logo_path):
    if logo_path:
        dimensions = calculate_logo_dimensions(logo_path)
        attributes = f' width="{dimensions[0]}" height="{dimensions[1]}"' if dimensions else ""
        return f'{signature_html}\n<p><img src="cid:signature_logo"{attributes}></p>'
    return signature_html


# ==================== STEP 6: BODY TEMPLATE ============================
# The template contains ONLY the specific body of the message (without
# greeting and without signature). The greeting "Dear <Name>," is
# generated automatically by the script from the spreadsheet's name
# column, and the signature is always the one extracted from the Word
# document in the previous step.

def extract_docx_text(path):
    return "\n".join(_extract_docx_paragraphs_html(path))


def extract_text_or_html_body(path):
    with open(path, "r", encoding="utf-8") as f:
        text = f.read()
    if path.lower().endswith((".html", ".htm")):
        return text
    lines = text.replace("\r\n", "\n").split("\n")
    return "\n".join(f"<p>{line.strip()}</p>" for line in lines if line.strip())


def wrap_in_html(text):
    if "<html" in text.lower():
        return text
    return (
        '<html>\n<body style="font-family:Calibri, Arial, sans-serif; font-size:11pt; color:#000000;">\n'
        f"{text}\n</body>\n</html>"
    )


def select_template(signature_html):
    print("=" * 70)
    print("STEP 6/6 - E-mail body template")
    print("=" * 70)
    print("The template must contain only the specific body of the message (no")
    print("initial greeting and no signature) — the greeting is generated")
    print("automatically from the name in the spreadsheet, and the signature is")
    print("always the one extracted from the Word document provided in the previous step.")
    while True:
        path = clean_path(input("Path to the e-mail body template (.txt, .html or .docx): "))
        if not os.path.exists(path) or not path.lower().endswith((".txt", ".html", ".htm", ".docx")):
            print("File not found or invalid format (use .txt/.html/.docx). Try again.")
            continue
        if os.path.getsize(path) == 0:
            print("The file is empty (0 bytes). Open it, add content and save again. Try another path.")
            continue
        try:
            if path.lower().endswith(".docx"):
                body = extract_docx_text(path)
            else:
                body = extract_text_or_html_body(path)
        except Exception as e:
            print(f"Could not read the file ({e}). Try again.")
            continue
        if not body.strip():
            print("The template is empty. Try again.")
            continue
        break

    # The body and the signature come from separate Word documents, each
    # with its own leading/trailing blank lines already stripped (see
    # _extract_docx_paragraphs_html). Without this explicit blank line
    # here, the last line of the body would be glued to the signature
    # (both use margin:0 to preserve Word's internal spacing), making the
    # signature look like it "starts" on its second visible line.
    full_body = f'<p>Dear {{name}},</p>\n{body}\n<p style="margin:0">&nbsp;</p>\n{signature_html}'
    final_html = wrap_in_html(full_body)
    print()
    return path, final_html


# ==================== SUMMARY AND CONFIRMATION ===============================

def count_valid_rows(df, name_column, email_column):
    valid = df[name_column].notna() & (df[name_column].astype(str).str.strip() != "") \
        & df[email_column].notna() & (df[email_column].astype(str).str.strip() != "")
    return int(valid.sum())


def ask_read_receipt():
    print("=" * 70)
    print("Read receipt")
    print("=" * 70)
    print("Outlook may ask the recipient for a read receipt (the recipient can")
    print("still decline sending it). After sending, use this script's 'Check read")
    print("receipts' option, pointing to the generated log, to update each")
    print("contact's status.")
    answer = input("Request a read receipt for this send? (y/n): ").strip().lower()
    return answer == "y"


def ask_generate_log(spreadsheet):
    print("=" * 70)
    print("Sending log")
    print("=" * 70)
    answer = input("Do you want to generate a sending log (Excel file) with the status of each e-mail? (y/n): ").strip().lower()
    if answer != "y":
        print()
        return False, None

    default_path = os.path.join(os.path.dirname(spreadsheet["path"]), "send_log.xlsx")
    path = clean_path(input(
        f"Path where the log should be saved (press Enter to use the default: {default_path}): "
    ))
    if not path:
        print()
        return True, default_path

    if os.path.isdir(path):
        path = os.path.join(path, "send_log.xlsx")
    elif not path.lower().endswith((".xlsx", ".xls")):
        path = path + ".xlsx"
    print()
    return True, path


def build_contact_table_rows(df, name_column, email_column):
    """Builds the (name, email) list of every contact detected in the
    spreadsheet (any row with a filled-in name), using 'N/A' for rows
    whose matching e-mail is missing/empty."""
    import pandas as pd

    rows = []
    for raw_name, raw_email in zip(df[name_column], df[email_column]):
        if pd.isna(raw_name):
            continue
        name = str(raw_name).strip()
        if not name:
            continue
        if pd.isna(raw_email) or not str(raw_email).strip():
            email = "N/A"
        else:
            email = str(raw_email).strip()
        rows.append((name, email))
    return rows


def print_contact_table(spreadsheet):
    rows = build_contact_table_rows(spreadsheet["df"], spreadsheet["name_column"], spreadsheet["email_column"])

    print("\n" + "=" * 70)
    print("NAMES AND E-MAILS DETECTED IN THE SPREADSHEET")
    print("=" * 70)
    if not rows:
        print("No contacts detected.")
    else:
        name_width = max(len("Name"), max(len(name) for name, _ in rows))
        email_width = max(len("Email"), max(len(email) for _, email in rows))
        print(f"{'Name'.ljust(name_width)}  {'Email'.ljust(email_width)}")
        print("-" * (name_width + email_width + 2))
        for name, email in rows:
            print(f"{name.ljust(name_width)}  {email.ljust(email_width)}")
    print("=" * 70 + "\n")


def show_summary_and_confirm(user_email, spreadsheet, sending_mailbox, logo_path, template_path, subject,
                              request_read_receipt, generate_log, log_path):
    valid_count = count_valid_rows(spreadsheet["df"], spreadsheet["name_column"], spreadsheet["email_column"])
    mailbox_label = sending_mailbox["sending_address"] or "default Outlook account"
    print("=" * 70)
    print("CONFIGURATION SUMMARY")
    print("=" * 70)
    print(f"Sent by (log):            {user_email}")
    print(f"Sending mailbox:          {mailbox_label}")
    print(f"Sending mode:             {'Send on behalf of (shared mailbox)' if sending_mailbox['mode'] == 'shared_mailbox' else 'Direct account'}")
    print(f"Subject:                  {subject}")
    print(f"Spreadsheet:              {spreadsheet['path']}")
    print(f"Sheet:                    {spreadsheet['sheet']}")
    print(f"Name column:              {spreadsheet['name_column']}")
    print(f"E-mail column:            {spreadsheet['email_column']}")
    print(f"Logo in the signature:    {logo_path or 'None'}")
    print(f"Template:                 {template_path}")
    print(f"Valid rows to send:       {valid_count}")
    print(f"Read receipt:             {'Yes' if request_read_receipt else 'No'}")
    print(f"Sending log:              {'Yes (' + log_path + ')' if generate_log else 'No'}")
    print("=" * 70)

    while True:
        print("\nWhat would you like to do?")
        print("  1. Proceed and send the e-mails")
        print("  2. Check the names and e-mails that will be used")
        print("  3. Cancel the sending")
        choice = input("Choose an option (1/2/3): ").strip()
        if choice == "1":
            return True
        if choice == "2":
            print_contact_table(spreadsheet)
            continue
        if choice == "3":
            return False
        print("Invalid option. Type 1, 2 or 3.")


# ==================== SENDING ===================================================

def send_emails(outlook, sending_mailbox, spreadsheet, logo_path, body_html_template, subject,
                 request_read_receipt=False, generate_log=True, log_path=None):
    df = spreadsheet["df"]
    name_column = spreadsheet["name_column"]
    email_column = spreadsheet["email_column"]

    import pandas as pd
    from datetime import datetime

    initial_read_status = "Pending" if request_read_receipt else "Not requested"

    log = []
    for raw_name, raw_email in zip(df[name_column], df[email_column]):
        if pd.isna(raw_name):
            continue
        name = str(raw_name).strip()
        if not name:
            continue

        if pd.isna(raw_email) or not str(raw_email).strip():
            log.append({
                "Name": name, "Email": "", "Subject": subject,
                "Status": "Skipped: no matching email found in contact list",
                "SentAt": "", "ReadReceipt": "N/A", "ReadAt": "",
            })
            print(f"[SKIPPED] {name}: no matching email found in contact list.")
            continue

        email = str(raw_email).strip()
        if "/" in name or "/" in email:
            log.append({
                "Name": name, "Email": email, "Subject": subject, "Status": "Skipped: multiple values in cell",
                "SentAt": "", "ReadReceipt": "N/A", "ReadAt": "",
            })
            print(f"[SKIPPED] {name} / {email}: cell with multiple values, needs to be split first.")
            continue

        try:
            mail = outlook.CreateItem(0)  # 0 = olMailItem
            mail.To = email
            mail.Subject = subject
            mail.HTMLBody = body_html_template.replace("{name}", name)

            if logo_path:
                attachment = mail.Attachments.Add(logo_path)
                attachment.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001E",
                    "signature_logo"
                )

            if sending_mailbox["account_obj"] is not None:
                mail.SendUsingAccount = sending_mailbox["account_obj"]
            elif sending_mailbox["mode"] == "shared_mailbox" and sending_mailbox["sending_address"]:
                mail.SentOnBehalfOfName = sending_mailbox["sending_address"]

            if request_read_receipt:
                mail.ReadReceiptRequested = True
            mail.Send()

            log.append({
                "Name": name, "Email": email, "Subject": subject, "Status": "Sent",
                "SentAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "ReadReceipt": initial_read_status, "ReadAt": "",
            })
            print(f"[OK] Sent to {name} ({email})")

        except Exception as e:
            log.append({
                "Name": name, "Email": email, "Subject": subject, "Status": f"Error: {e}",
                "SentAt": "", "ReadReceipt": "N/A", "ReadAt": "",
            })
            print(f"[ERROR] Failed to send to {email}: {e}")

        time.sleep(2)

    if not generate_log:
        print("\nDone. Sending log generation disabled for this run.")
        return None

    log_df = pd.DataFrame(log)
    log_path = log_path or os.path.join(os.path.dirname(spreadsheet["path"]), "send_log.xlsx")
    log_df.to_excel(log_path, index=False)
    print(f"\nDone. Log saved to: {log_path}")
    if request_read_receipt:
        print("Read receipts requested. Use this script's 'Check read receipts' option,")
        print(f"pointing to the file above, to update the status.")
    return log_path


# ==================== READ RECEIPT CHECK =====================================
# When ReadReceiptRequested=True is set on sending, Outlook makes the
# recipient receive a read receipt request; if they accept it, a report
# item (MessageClass "REPORT.IPM.Note.IPNRN" for confirmed or
# "REPORT.IPM.Note.IPNNRN" for declined — some setups may omit the
# "REPORT." prefix) arrives in the Inbox of the
# account that sent the message. This section scans that inbox and
# cross-references the found receipts with the sending log (by recipient
# e-mail + subject), updating the read status of each row.

def locate_shared_mailbox_inbox(namespace, target):
    """Searches, among the mailboxes available in the current Outlook
    profile, for the one matching the given e-mail/display name and
    returns its "Inbox" folder. Returns None if not found (e.g., the
    shared mailbox was not added to the profile)."""
    inbox_names = ["Inbox", "Caixa de Entrada"]
    target_lower = target.strip().lower()

    for root_folder in namespace.Folders:
        try:
            root_name = (root_folder.Name or "").strip().lower()
        except Exception:
            continue
        if target_lower in root_name:
            for inbox_name in inbox_names:
                try:
                    return root_folder.Folders(inbox_name)
                except Exception:
                    continue
    return None


def extract_original_subject_from_receipt(receipt_subject):
    """Removes common prefixes (en/pt) that Outlook adds to receipt report
    subjects, returning the original subject of the sent e-mail."""
    prefixes = ["read:", "lido:", "declined:", "recusado:", "delivered:", "entregue:"]
    text = (receipt_subject or "").strip()
    lower = text.lower()
    for prefix in prefixes:
        if lower.startswith(prefix):
            return text[len(prefix):].strip()
    return text


def ask_mailbox_for_receipt_check(namespace):
    """Reuses the same Outlook profile mailbox detection (full accounts and
    shared mailboxes/additional folders) used in the sending mailbox
    selection step, so the user can consistently choose which mailbox to
    check for read receipts."""
    mailboxes = list_available_mailboxes(namespace)

    print("Mailboxes detected in the Outlook profile:")
    for i, mailbox in enumerate(mailboxes, start=1):
        type_label = "full account" if mailbox["type"] == "account" else "additional folder / possible shared mailbox"
        print(f"  {i}. {mailbox['label']} ({type_label})")

    while True:
        choice = input(
            "E-mail or display name of the mailbox used in the original send "
            "(list number above or type the e-mail/name): "
        ).strip()
        if choice.isdigit() and 1 <= int(choice) <= len(mailboxes):
            return mailboxes[int(choice) - 1]["label"]
        if choice:
            return choice
        print("Empty value. Type the number of a list option or a valid e-mail/name.")


def get_receipt_smtp_address(item):
    """For internal Exchange contacts, the receipt item's SenderEmailAddress
    sometimes comes in X.500 (legacyExchangeDN) format instead of the SMTP
    address, which would break the comparison against the spreadsheet/log.
    This function tries a few ways to get the real SMTP address."""
    try:
        raw = str(getattr(item, "SenderEmailAddress", "") or "").strip()
    except Exception:
        raw = ""
    if "@" in raw:
        return raw.lower()

    try:
        smtp = item.PropertyAccessor.GetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
        )
        if smtp and "@" in smtp:
            return smtp.strip().lower()
    except Exception:
        pass

    try:
        sender = item.Sender
        if sender is not None and sender.AddressEntry.Type == "EX":
            smtp = sender.AddressEntry.GetExchangeUser().PrimarySmtpAddress
            if smtp and "@" in smtp:
                return smtp.strip().lower()
    except Exception:
        pass

    return raw.lower()


def check_read_receipts(namespace):
    print("=" * 70)
    print("READ RECEIPT CHECK")
    print("=" * 70)
    import pandas as pd

    while True:
        log_path = clean_path(input("Path to the sending log (send_log.xlsx): "))
        if os.path.exists(log_path) and log_path.lower().endswith((".xlsx", ".xls")):
            break
        print("File not found or invalid format (use .xlsx/.xls). Try again.")

    try:
        log_df = pd.read_excel(log_path)
    except Exception as e:
        print(f"[ERROR] Could not read the log ({e}).")
        return

    expected_columns = {"Name", "Email", "Status"}
    if not expected_columns.issubset(set(log_df.columns)):
        print(f"[ERROR] The log must contain the columns {expected_columns}. "
              f"Columns found: {list(log_df.columns)}")
        return

    if "Subject" not in log_df.columns:
        log_df["Subject"] = ""
    if "ReadReceipt" not in log_df.columns:
        log_df["ReadReceipt"] = "Pending"
    if "ReadAt" not in log_df.columns:
        log_df["ReadAt"] = ""

    # If the log already existed and these columns were completely
    # empty/blank (e.g., generated by a previous run of this script before
    # any receipts existed), pandas/Excel may have read them back as a
    # numeric dtype (float64, an all-NaN column). Force them to text so
    # values like "Confirmed"/date-time can be written without a type error.
    log_df["Subject"] = log_df["Subject"].astype(object).where(log_df["Subject"].notna(), "")
    log_df["ReadReceipt"] = log_df["ReadReceipt"].astype(object).where(
        log_df["ReadReceipt"].notna(), "Pending"
    )
    log_df["ReadAt"] = log_df["ReadAt"].astype(object).where(log_df["ReadAt"].notna(), "")

    target = ask_mailbox_for_receipt_check(namespace)

    print(f"\nLooking for the Inbox of '{target}' in the Outlook profile...")
    target_inbox_folder = locate_shared_mailbox_inbox(namespace, target)
    default_inbox_folder = namespace.GetDefaultFolder(6)  # olFolderInbox

    folders_to_search = []
    if target_inbox_folder is None:
        print(f"[WARNING] '{target}' was not found as an additional mailbox in this profile.")
        print("        Checking only the current user's default Inbox.")
    else:
        print("Sending account's inbox located.")
        folders_to_search.append(target_inbox_folder)

    # When the original send was done "on behalf of" a shared mailbox
    # (SentOnBehalfOfName), Exchange usually delivers read receipts to the
    # Inbox of the personal account that actually sent the message, not to
    # the shared mailbox's own Inbox. So we always also check the current
    # user's default Inbox, in addition to the chosen mailbox (avoiding
    # duplicates if they turn out to be the same folder).
    if target_inbox_folder is None or target_inbox_folder.EntryID != default_inbox_folder.EntryID:
        folders_to_search.append(default_inbox_folder)

    # Instead of using Items.Restrict() with a DASL query over MessageClass
    # (which proved unreliable depending on the Outlook version/setup), we
    # walk every item in the folder and check MessageClass directly in
    # Python — slower, but much more reliable for finding read/decline
    # receipts.
    receipt_items = []
    seen_ids = set()
    unrecognized_candidates = []
    receipt_subject_prefixes = ("read:", "declined:", "delivered:", "lido:", "recusado:", "entregue:")
    for folder in folders_to_search:
        try:
            folder_items = folder.Items
            total_items = folder_items.Count
        except Exception as e:
            print(f"[ERROR] Failed to access items in '{folder.Name}': {e}")
            continue
        print(f"Checking {total_items} item(s) in '{folder.Name}'...")
        for item in folder_items:
            try:
                item_class = str(item.MessageClass)
            except Exception:
                continue
            if not (item_class.startswith("REPORT.IPM.Note.IPNRN") or item_class.startswith("REPORT.IPM.Note.IPNNRN")
                    or item_class.startswith("IPM.Note.IPNRN") or item_class.startswith("IPM.Note.IPNNRN")):
                # Diagnostics: if the subject looks like a receipt (e.g.
                # "Read:"/"Declined:" prefix) but the MessageClass didn't
                # match what we expect, keep it to show the user — helps
                # quickly spot if Outlook is using a different class than
                # documented.
                try:
                    item_subject = str(item.Subject or "")
                except Exception:
                    item_subject = ""
                if item_subject.strip().lower().startswith(receipt_subject_prefixes):
                    unrecognized_candidates.append((item_subject, item_class))
                continue
            try:
                entry_id = item.EntryID
            except Exception:
                entry_id = None
            if entry_id is not None and entry_id in seen_ids:
                continue
            if entry_id is not None:
                seen_ids.add(entry_id)
            receipt_items.append(item)

    if not receipt_items:
        print("No read receipts found in the checked mailboxes.")
        if unrecognized_candidates:
            print("\n[DIAGNOSTICS] Found item(s) with a receipt-like subject, but with a")
            print("MessageClass different from expected (IPM.Note.IPNRN/IPNNRN):")
            for item_subject, item_class in unrecognized_candidates:
                print(f"  - Subject: {item_subject!r} | MessageClass: {item_class!r}")

    # Instead of recomputing (and rescanning) the log's Email/Subject
    # columns for every single receipt, build a one-time index here: a
    # single O(log_rows) pass that lets each receipt do an O(1) lookup
    # instead of an O(log_rows) scan — matters when the log is large and
    # there are many receipts to cross-reference.
    index_by_email_and_subject = defaultdict(list)
    normalized_email = log_df["Email"].astype(str).str.strip().str.lower()
    normalized_subject = log_df["Subject"].astype(str).str.strip().str.lower()
    normalized_status = log_df["Status"].astype(str).str.strip()
    for idx in log_df.index[normalized_status == "Sent"]:
        email_idx = normalized_email[idx]
        subject_idx = normalized_subject[idx]
        index_by_email_and_subject[(email_idx, subject_idx)].append(idx)

    confirmed_count = 0
    declined_count = 0
    for item in receipt_items:
        try:
            msg_class = str(item.MessageClass)
            reader_email = get_receipt_smtp_address(item)
            original_subject = extract_original_subject_from_receipt(str(item.Subject or ""))
        except Exception as e:
            print(f"[DIAGNOSTICS] Failed to process a receipt (item skipped): {e}")
            continue

        # Report items (read/decline receipts) sometimes don't expose
        # ReceivedTime the same way a regular e-mail does; fall back to
        # CreationTime and, at worst, leave it blank instead of discarding
        # the whole receipt. The value is formatted as DD/MM/YYYY HH:MM for
        # readability in the log.
        try:
            read_date = item.ReceivedTime.strftime("%d/%m/%Y %H:%M")
        except Exception:
            try:
                read_date = item.CreationTime.strftime("%d/%m/%Y %H:%M")
            except Exception:
                read_date = ""

        declined = msg_class.startswith("REPORT.IPM.Note.IPNNRN") or msg_class.startswith("IPM.Note.IPNNRN")

        matching_indices = index_by_email_and_subject.get(
            (reader_email, original_subject.strip().lower()), []
        )
        if not matching_indices:
            # Fallback: only kicks in when the log row itself has no subject
            # recorded (e.g., a log generated by an older run of this script,
            # before the "Subject" column existed). In that case, match by
            # e-mail only within this log. Do NOT use this fallback when the
            # log row already has a different subject than the receipt: that
            # would let a receipt from an earlier send (e.g. "test 7") be
            # incorrectly attributed to a different send to the same
            # recipient (e.g. "test 9"), inflating the confirmation count.
            matching_indices = index_by_email_and_subject.get((reader_email, ""), [])
        if not matching_indices:
            print(
                f"[DIAGNOSTICS] Receipt received from '{reader_email}' (receipt subject: "
                f"'{original_subject}') did not match any row in the log (with Status='Sent')."
            )
            continue

        new_status = "Declined" if declined else "Confirmed"
        log_df.loc[matching_indices, "ReadReceipt"] = new_status
        log_df.loc[matching_indices, "ReadAt"] = read_date
        if declined:
            declined_count += 1
        else:
            confirmed_count += 1

    log_df.to_excel(log_path, index=False)

    sent_count = int((log_df["Status"] == "Sent").sum())
    pending_count = int(
        ((log_df["Status"] == "Sent") & (log_df["ReadReceipt"] == "Pending")).sum()
    )
    print("\n" + "=" * 70)
    print("CHECK RESULT")
    print("=" * 70)
    print(f"E-mails sent in the log:              {sent_count}")
    print(f"Read receipts received:               {confirmed_count}")
    print(f"Receipt declines received:            {declined_count}")
    print(f"Still pending (no response):          {pending_count}")
    print(f"\nUpdated log saved to: {log_path}")


# ==================== ORCHESTRATION ============================================

def choose_operation_mode():
    print("=" * 70)
    print("What do you want to do?")
    print("=" * 70)
    print("  1. New e-mail send")
    print("  2. Check read receipts for a send already done")
    while True:
        choice = input("Choose an option (1/2): ").strip()
        if choice in ("1", "2"):
            return choice
        print("Invalid option. Type 1 or 2.")


def main():
    mode = choose_operation_mode()
    print()

    run_environment_check()

    import win32com.client as win32
    outlook = win32.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    if mode == "2":
        check_read_receipts(namespace)
        return

    available_mailboxes = list_available_mailboxes(namespace)
    user_email = ask_user_email(available_mailboxes)
    sending_mailbox = ask_sending_mailbox(namespace, available_mailboxes)
    spreadsheet = select_spreadsheet()
    _, signature_docx_html = ask_signature_path()
    logo_path = ask_logo(signature_docx_html)
    signature_html = build_final_signature(signature_docx_html, logo_path)
    template_path, body_html_template = select_template(signature_html)

    subject = input("E-mail subject: ").strip()
    request_read_receipt = ask_read_receipt()
    generate_log, log_path = ask_generate_log(spreadsheet)

    if not show_summary_and_confirm(user_email, spreadsheet, sending_mailbox, logo_path, template_path, subject,
                                     request_read_receipt, generate_log, log_path):
        print("Operation cancelled by the user.")
        return

    send_emails(outlook, sending_mailbox, spreadsheet, logo_path, body_html_template, subject,
                request_read_receipt, generate_log, log_path)


if __name__ == "__main__":
    main()
