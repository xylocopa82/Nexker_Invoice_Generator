#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import sys
import json
import webbrowser
import subprocess
import urllib.parse
import smtplib
import threading
from datetime import datetime, timedelta

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

try:
    from fpdf import FPDF
except ImportError:
    print("Please install the PDF library first:   pip install fpdf2")
    sys.exit(1)

# =====================================================================
#  BUSINESS CONSTANTS
# =====================================================================
BUSINESS = {
    "name": "NEXKER Web & Digital Solutions",
    "website": "www.nexker.com",
    "phone": "00 91 8281829736",
    "email": "nexkerwebsolutions@gmail.com",
    "instagram": "instagram.com/nexke_r",
    "facebook": "facebook.com/nexkerwebsolutions",
}

WEBSITE_SERVICES = [
    "Business / Corporate Website", "Portfolio Website", "E-commerce Website",
    "Blog / Publishing Website", "News Website", "Educational / E-learning Website",
    "LMS, Learning Management System", "Booking / Reservation Website",
    "Marketplace Website", "Social Networking Website", "Community / Forum Website",
    "SaaS Website / Web Application", "CRM Website", "ERP Website", "Job Portal",
    "Directory Website", "Real Estate Website", "NGO / Non-profit Website",
    "Government / Public Service Website", "Landing Page", "Membership Website",
    "Crowdfunding Website", "Booking + Service Marketplace", "Web Portal",
]

MARKETING_SERVICES = [
    "Social Media Management - Facebook",
    "Social Media Management - Instagram",
    "Social Media Management - Google My Business Profile",
    "Poster Design",
    "Local SEO",
]

ALL_SERVICES = ["Custom / Other Service"] + WEBSITE_SERVICES + MARKETING_SERVICES

CURRENCIES = {
    "INR": {"symbol": "Rs.", "tax_label": "GST", "default_tax": 18},
    "AED": {"symbol": "AED", "tax_label": "VAT", "default_tax": 5},
    "USD": {"symbol": "$",   "tax_label": "Tax", "default_tax": 0},
}

# Files live next to the program (works for both .py and .exe)
if getattr(sys, "frozen", False):                 # running as packaged .exe
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

SETTINGS_FILE = os.path.join(APP_DIR, "nexker_settings.json")
TOKEN_FILE    = os.path.join(APP_DIR, "token.json")
CREDS_FILE    = os.path.join(APP_DIR, "credentials.json")

DEFAULTS = {
    "invoice_counter": 1,
    "save_folder": os.path.join(os.path.expanduser("~"), "Documents", "NEXKER Invoices"),
    "drive_folder_id": "",
    "last_currency": "INR",
    "email_app_password": "",
    "gstin": "",
    "default_notes": "Thank you for your business!",
    "default_terms": ("1. Payment due within 7 days of the invoice date.\n"
                      "2. 50% advance is required to start the project.\n"
                      "3. Prices once agreed are final.\n"
                      "4. Third-party costs (domain, hosting, ads) are billed separately.\n"
                      "5. Support/warranty as per the agreed project scope."),
}

NAVY = (16, 36, 84)
ACCENT = (0, 174, 239)
GRAY = (105, 115, 130)
DARK = (35, 45, 65)
LIGHT = (240, 244, 250)
MID = (222, 230, 242)
NAVY_HEX = "#10244f"

# =====================================================================
#  SMALL HELPERS
# =====================================================================
def load_settings():
    s = dict(DEFAULTS)
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                s.update(json.load(f))
        except Exception:
            pass
    return s


def save_settings(s):
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(s, f, indent=2)
    except Exception:
        pass


def sanitize_filename(s):
    return re.sub(r'[^\w\-. ]+', '_', s).strip() or "Client"


def slug(s):
    return re.sub(r"\W+", "_", s.lower()).strip("_")


def fnum(v):
    return ("%g" % v)


def fmt_money(v, cur):
    return "{} {:,.2f}".format(CURRENCIES[cur]["symbol"], v)


def asc(s):
    """Make text safe for PDF core fonts (latin-1)."""
    s = str(s)
    for a, b in (("\u2013", "-"), ("\u2014", "-"), ("\u2018", "'"), ("\u2019", "'"),
                 ("\u201c", '"'), ("\u201d", '"'), ("\u20b9", "Rs."), ("\u2022", "-"),
                 ("\u00a0", " ")):
        s = s.replace(a, b)
    return s.encode("latin-1", "replace").decode("latin-1")


def wrapped_lines(pdf, text, width):
    """Word-wrap text for the current pdf font; returns list of lines."""
    lines = []
    for para in asc(text).split("\n"):
        words = para.split()
        if not words:
            lines.append("")
            continue
        cur = words[0]
        for wd in words[1:]:
            test = cur + " " + wd
            if pdf.get_string_width(test) <= width:
                cur = test
            else:
                lines.append(cur)
                cur = wd
        lines.append(cur)
    return lines


def amount_in_words(amount, cur):
    whole = int(amount)
    cents = int(round((amount - whole) * 100))
    if cents == 100:
        whole += 1
        cents = 0

    ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
            "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
            "Seventeen", "Eighteen", "Nineteen"]
    tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]

    def under1000(n):
        out = []
        if n >= 100:
            out.append(ones[n // 100] + " Hundred")
            n %= 100
        if n >= 20:
            t = tens[n // 10]
            if n % 10:
                t += "-" + ones[n % 10]
            out.append(t)
        elif n:
            out.append(ones[n])
        return " ".join(out)

    def indian(n):
        out = []
        for value, name in ((10000000, "Crore"), (100000, "Lakh"), (1000, "Thousand")):
            if n >= value:
                out.append(under1000(n // value) + " " + name)
                n %= value
        if n:
            out.append(under1000(n))
        return " ".join(out)

    def intl(n):
        out = []
        for value, name in ((1000000000, "Billion"), (1000000, "Million"), (1000, "Thousand")):
            if n >= value:
                out.append(under1000(n // value) + " " + name)
                n %= value
        if n:
            out.append(under1000(n))
        return " ".join(out)

    words = "Zero" if (whole == 0 and cents == 0) else (indian(whole) if cur == "INR" else intl(whole))
    big, small = {"INR": ("Rupees", "Paise"), "AED": ("UAE Dirhams", "Fils"),
                  "USD": ("US Dollars", "Cents")}[cur]
    return "{} {} and {:02d} {} Only".format(big, words, cents, small)


def open_file(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


def reveal_file(path):
    """Open the folder with the file selected (for drag-drop into WhatsApp)."""
    try:
        if sys.platform.startswith("win"):
            subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
        elif sys.platform == "darwin":
            subprocess.Popen(["open", "-R", path])
        else:
            subprocess.Popen(["xdg-open", os.path.dirname(path)])
    except Exception:
        pass

# =====================================================================
#  PDF CREATION
# =====================================================================
class _InvPDF(FPDF):
    def footer(self):
        self.set_y(-18)
        self.set_draw_color(*MID)
        self.set_line_width(0.3)
        self.line(10, self.get_y(), 200, self.get_y())
        self.set_y(-16)
        self.set_font("Helvetica", "", 7.5)
        self.set_text_color(*GRAY)
        self.cell(0, 4, "Thank you for your business!", align="C", new_x="LMARGIN", new_y="NEXT")
        self.cell(0, 4, "www.nexker.com   |   instagram.com/nexke_r   |   facebook.com/nexkerwebsolutions",
                  align="C", new_x="LMARGIN", new_y="NEXT")
        self.set_y(-8)
        self.cell(0, 4, "Page %d" % self.page_no(), align="C")


def build_invoice_pdf(d, path):
    pdf = _InvPDF()
    pdf.set_margins(10, 10, 10)
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.add_page()
    cur = d["currency"]

    # ---- totals ----
    sub = sum(i["qty"] * i["rate"] for i in d["items"])
    disc = max(float(d.get("disc", 0) or 0), 0)
    tax_pct = float(d.get("tax_pct", 0) or 0)
    adv = max(float(d.get("adv", 0) or 0), 0)
    taxable = max(sub - disc, 0)
    tax = taxable * tax_pct / 100.0
    total = round(taxable + tax, 2)
    balance = round(max(total - adv, 0), 2)
    tax_label = CURRENCIES[cur]["tax_label"]

    # ---- header band ----
    pdf.set_fill_color(*NAVY)
    pdf.rect(0, 0, 210, 32, style="F")
    pdf.set_fill_color(*ACCENT)
    pdf.rect(0, 32, 210, 1.4, style="F")

    pdf.set_xy(12, 7)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 21)
    pdf.cell(w=0, h=9, text="NEXKER")
    pdf.set_xy(12, 16.5)
    pdf.set_font("Helvetica", "", 10.5)
    pdf.set_text_color(190, 214, 255)
    pdf.cell(w=0, h=5, text="Web & Digital Solutions")
    if d.get("tax_id"):
        pdf.set_xy(12, 23)
        pdf.set_font("Helvetica", "", 8)
        pdf.cell(w=0, h=4, text=asc(d["tax_id"]))

    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(230, 238, 250)
    for i, txt in enumerate([BUSINESS["website"], BUSINESS["email"], BUSINESS["phone"]]):
        pdf.set_xy(105, 8 + i * 5.5)
        pdf.cell(w=93, h=5.5, text=txt, align="R")

    # ---- INVOICE title + meta ----
    pdf.set_text_color(*NAVY)
    pdf.set_xy(10, 40)
    pdf.set_font("Helvetica", "B", 25)
    pdf.cell(w=0, h=11, text="INVOICE")

    y = 40
    for label, val, bold in [("Invoice No", d["inv_no"], True),
                             ("Invoice Date", d["date_str"], False),
                             ("Due Date", d["due_str"], False),
                             ("Currency", cur, False)]:
        pdf.set_font("Helvetica", "", 9.5)
        pdf.set_text_color(*GRAY)
        pdf.set_xy(128, y)
        pdf.cell(w=40, h=5.5, text=label + ":", align="R")
        pdf.set_font("Helvetica", "B", 9.5)
        pdf.set_text_color(*DARK)
        pdf.set_xy(169, y)
        pdf.cell(w=29, h=5.5, text=asc(val), align="R")
        y += 6

    # ---- Bill To ----
    y = 62
    pdf.set_xy(10, y)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_text_color(*ACCENT)
    pdf.cell(w=0, h=5, text="BILL  TO")
    y += 6
    pdf.set_xy(10, y)
    pdf.set_font("Helvetica", "B", 12.5)
    pdf.set_text_color(*DARK)
    pdf.cell(w=95, h=6, text=asc(d["client_name"]))
    y += 7
    if d["client_company"]:
        pdf.set_xy(10, y)
        pdf.set_font("Helvetica", "", 9.5)
        pdf.cell(w=95, h=5, text=asc(d["client_company"]))
        y += 5.5
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(*GRAY)
    for ln in wrapped_lines(pdf, d["client_address"], 95)[:3]:
        pdf.set_xy(10, y)
        pdf.cell(w=95, h=4.6, text=ln)
        y += 4.6
    contact = "  |  ".join(x for x in [d["client_phone"], d["client_email"]] if x)
    if contact:
        pdf.set_xy(10, y)
        pdf.cell(w=95, h=4.6, text=asc(contact))
        y += 4.6

    # ---- Total-due box ----
    pdf.set_fill_color(*LIGHT)
    pdf.rect(128, 62, 70, 22, style="F")
    pdf.set_draw_color(*ACCENT)
    pdf.set_line_width(0.5)
    pdf.rect(128, 62, 70, 22)
    pdf.set_xy(132, 64.5)
    pdf.set_font("Helvetica", "B", 8)
    pdf.set_text_color(*GRAY)
    pdf.cell(w=0, h=4, text="BALANCE DUE" if adv > 0 else "TOTAL DUE")
    pdf.set_xy(132, 69)
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(*NAVY)
    pdf.cell(w=0, h=7, text=fmt_money(balance, cur))
    pdf.set_xy(132, 77.5)
    pdf.set_font("Helvetica", "", 7.5)
    pdf.set_text_color(*GRAY)
    pdf.cell(w=0, h=3.5, text="Due Date: " + asc(d["due_str"]))

    # ---- items table ----
    def table_header(y):
        pdf.set_fill_color(*NAVY)
        pdf.rect(10, y, 180, 7.5, style="F")
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.set_text_color(255, 255, 255)
        pdf.set_xy(10, y);  pdf.cell(w=8,  h=7.5, text="#")
        pdf.set_xy(18, y);  pdf.cell(w=100, h=7.5, text="SERVICE  /  DESCRIPTION")
        pdf.set_xy(118, y); pdf.cell(w=16, h=7.5, text="QTY", align="C")
        pdf.set_xy(134, y); pdf.cell(w=26, h=7.5, text="RATE", align="R")
        pdf.set_xy(160, y); pdf.cell(w=30, h=7.5, text="AMOUNT", align="R")

    y = max(y + 5, 92)
    table_header(y)
    y += 7.5

    for i, it in enumerate(d["items"], 1):
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_text_color(*DARK)
        svc_lines = wrapped_lines(pdf, it["service"], 96)
        pdf.set_font("Helvetica", "", 8.3)
        pdf.set_text_color(*GRAY)
        dsc_lines = wrapped_lines(pdf, it["desc"], 96) if it["desc"] else []
        h = max(8, len(svc_lines) * 4.8 + len(dsc_lines) * 4.2 + 4)

        if y + h > 262:                      # new page
            pdf.add_page()
            pdf.set_fill_color(*NAVY)
            pdf.rect(0, 0, 210, 10, style="F")
            pdf.set_xy(10, 2.5)
            pdf.set_font("Helvetica", "B", 10)
            pdf.set_text_color(255, 255, 255)
            pdf.cell(w=0, h=5, text="NEXKER Web & Digital Solutions  -  Invoice " + asc(d["inv_no"]))
            y = 16
            table_header(y)
            y += 7.5

        if i % 2 == 0:
            pdf.set_fill_color(*LIGHT)
            pdf.rect(10, y, 180, h, style="F")

        yy = y + 2
        pdf.set_xy(10, yy)
        pdf.set_font("Helvetica", "", 8.5)
        pdf.set_text_color(*GRAY)
        pdf.cell(w=8, h=4.5, text=str(i))
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_text_color(*DARK)
        for ln in svc_lines:
            pdf.set_xy(18, yy); pdf.cell(w=96, h=4.8, text=ln); yy += 4.8
        pdf.set_font("Helvetica", "", 8.3)
        pdf.set_text_color(*GRAY)
        for ln in dsc_lines:
            pdf.set_xy(18, yy); pdf.cell(w=96, h=4.2, text=ln); yy += 4.2
        pdf.set_xy(118, y + 2)
        pdf.set_font("Helvetica", "", 9)
        pdf.set_text_color(*DARK)
        pdf.cell(w=16, h=4.5, text=fnum(it["qty"]), align="C")
        pdf.set_xy(134, y + 2)
        pdf.cell(w=26, h=4.5, text=fmt_money(it["rate"], cur), align="R")
        pdf.set_xy(160, y + 2)
        pdf.set_font("Helvetica", "B", 9)
        pdf.cell(w=30, h=4.5, text=fmt_money(it["qty"] * it["rate"], cur), align="R")
        pdf.set_draw_color(*MID)
        pdf.set_line_width(0.2)
        pdf.line(10, y + h, 190, y + h)
        y += h

    # ---- totals ----
    y += 8

    def line(lbl, val, bold=False):
        nonlocal y
        pdf.set_font("Helvetica", "B" if bold else "", 9.5)
        pdf.set_text_color(*(NAVY if bold else GRAY))
        pdf.set_xy(120, y); pdf.cell(w=40, h=5.5, text=lbl, align="R")
        pdf.set_text_color(*DARK)
        pdf.set_xy(160, y); pdf.cell(w=30, h=5.5, text=val, align="R")
        y += 6

    line("Subtotal", fmt_money(sub, cur))
    if disc > 0:
        line("Discount  (-)", fmt_money(disc, cur))
    if tax_pct > 0:
        line("{} @ {}%".format(tax_label, fnum(tax_pct)), fmt_money(tax, cur))

    pdf.set_fill_color(*NAVY)
    pdf.rect(112, y + 1, 78, 10, style="F")
    pdf.set_xy(116, y + 1)
    pdf.set_font("Helvetica", "B", 10.5)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(w=40, h=10, text="TOTAL")
    pdf.set_xy(146, y + 1)
    pdf.cell(w=40, h=10, text=fmt_money(total, cur), align="R")
    y += 14
    if adv > 0:
        line("Advance Received  (-)", fmt_money(adv, cur))
        line("BALANCE DUE", fmt_money(balance, cur), bold=True)

    # ---- amount in words ----
    y += 4
    pdf.set_xy(10, y)
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.set_text_color(*GRAY)
    pdf.cell(w=0, h=4, text="Amount in Words:")
    y += 4.5
    pdf.set_font("Helvetica", "BI", 8.8)
    pdf.set_text_color(*DARK)
    for ln in wrapped_lines(pdf, amount_in_words(total, cur), 180):
        pdf.set_xy(10, y)
        pdf.cell(w=180, h=4.4, text=ln)
        y += 4.4

    # ---- payment details ----
    bank = d.get("bank", {})
    blines = []
    if bank.get("Bank Name"):     blines.append(("Bank",         bank["Bank Name"]))
    if bank.get("A/C Holder"):    blines.append(("Account Name", bank["A/C Holder"]))
    if bank.get("A/C Number"):    blines.append(("Account No",   bank["A/C Number"]))
    if bank.get("IFSC / SWIFT"):  blines.append(("IFSC / SWIFT", bank["IFSC / SWIFT"]))
    if bank.get("UPI ID"):        blines.append(("UPI",          bank["UPI ID"]))
    if blines:
        y += 4
        bh = 6 + len(blines) * 5
        pdf.set_fill_color(*LIGHT)
        pdf.rect(10, y, 105, bh, style="F")
        pdf.set_xy(14, y + 1.5)
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.set_text_color(*NAVY)
        pdf.cell(w=0, h=4.5, text="PAYMENT DETAILS")
        yy = y + 7
        for k, v in blines:
            pdf.set_xy(14, yy)
            pdf.set_font("Helvetica", "", 8.3)
            pdf.set_text_color(*GRAY)
            pdf.cell(w=30, h=4.6, text=k)
            pdf.set_xy(46, yy)
            pdf.set_font("Helvetica", "B", 8.3)
            pdf.set_text_color(*DARK)
            pdf.cell(w=66, h=4.6, text=asc(v))
            yy += 5
        y += bh

    # ---- notes / terms ----
    def block(title, text):
        nonlocal y
        if not text.strip():
            return
        if y > 235:
            pdf.add_page()
            y = 16
        y += 5
        pdf.set_xy(10, y)
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.set_text_color(*NAVY)
        pdf.cell(w=0, h=4.5, text=title)
        y += 5.5
        pdf.set_font("Helvetica", "", 8)
        pdf.set_text_color(*GRAY)
        for ln in wrapped_lines(pdf, text, 180):
            if y > 275:
                pdf.add_page()
                y = 16
            pdf.set_xy(10, y)
            pdf.cell(w=180, h=4.1, text=ln)
            y += 4.1

    block("NOTES", d.get("notes", ""))
    block("TERMS & CONDITIONS", d.get("terms", ""))

    pdf.output(path)

# =====================================================================
#  GOOGLE DRIVE HELPERS
# =====================================================================
def _drive_service():
    try:
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request
        from google_auth_oauthlib.flow import InstalledAppFlow
        from googleapiclient.discovery import build
    except ImportError:
        raise RuntimeError("Google libraries are not installed.\n\nRun:\n"
                           "  pip install google-api-python-client google-auth-oauthlib\n\n"
                           "See Help > Setup Guide.")
    scopes = ["https://www.googleapis.com/auth/drive.file"]
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, scopes)
    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except Exception:
            creds = None
    if not creds:
        if not os.path.exists(CREDS_FILE):
            raise RuntimeError("credentials.json not found next to the program.\n\n"
                               "See Help > Setup Guide for the one-time Google Drive setup.")
        flow = InstalledAppFlow.from_client_secrets_file(CREDS_FILE, scopes)
        creds = flow.run_local_server(port=0)
    with open(TOKEN_FILE, "w") as tf:
        tf.write(creds.to_json())
    return build("drive", "v3", credentials=creds)


def upload_to_gdrive(file_path, folder_id=""):
    svc = _drive_service()
    from googleapiclient.http import MediaFileUpload
    meta = {"name": os.path.basename(file_path)}
    if folder_id:
        meta["parents"] = [folder_id]
    media = MediaFileUpload(file_path, mimetype="application/pdf")
    f = svc.files().create(body=meta, media_body=media, fields="id, webViewLink").execute()
    return f.get("webViewLink", "Uploaded (ID: %s)" % f.get("id"))

# =====================================================================
#  MAIN APPLICATION
# =====================================================================
class InvoiceApp:
    def __init__(self, root):
        self.root = root
        self.settings = load_settings()
        self.items = []
        self.pdf_path = None
        self.totals = {"total": 0}

        root.title("NEXKER Web & Digital Solutions - Invoice Generator")
        root.geometry("1250x830")
        root.minsize(1150, 780)

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(".", font=("Helvetica", 9))
        style.configure("TLabelframe.Label", font=("Helvetica", 9, "bold"), foreground=NAVY_HEX)

        self._auto_inv = self._next_inv_no()
        self._build_menu()
        self._build_ui()
        self._load_form_defaults()
        self._recalc()
        root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------- menu -------------------------
    def _build_menu(self):
        m = tk.Menu(self.root)
        f = tk.Menu(m, tearoff=0)
        f.add_command(label="New Invoice", command=self._new_invoice, accelerator="Ctrl+N")
        f.add_separator()
        f.add_command(label="Choose Local Save Folder...", command=self._choose_folder)
        f.add_separator()
        f.add_command(label="Exit", command=self._on_close)
        m.add_cascade(label="File", menu=f)

        d = tk.Menu(m, tearoff=0)
        d.add_command(label="Choose Google Drive Folder...", command=self._pick_drive_folder)
        d.add_command(label="Clear Drive Login (token.json)", command=self._clear_token)
        m.add_cascade(label="Google Drive", menu=d)

        h = tk.Menu(m, tearoff=0)
        h.add_command(label="Setup Guide (Drive / Gmail / WhatsApp)", command=self._show_guide)
        h.add_command(label="About", command=self._about)
        m.add_cascade(label="Help", menu=h)
        self.root.config(menu=m)
        self.root.bind("<Control-n>", lambda e: self._new_invoice())
        self.root.bind("<Control-g>", self.generate)

    # ------------------------- UI -------------------------
    def _build_ui(self):
        wrap = tk.Frame(self.root, bg="#eef1f6")
        wrap.pack(fill="both", expand=True)

        banner = tk.Frame(wrap, bg=NAVY_HEX)
        banner.pack(fill="x")
        tk.Label(banner, text="NEXKER", font=("Helvetica", 16, "bold"),
                 bg=NAVY_HEX, fg="white").pack(side="left", padx=(14, 4), pady=8)
        tk.Label(banner, text="Web & Digital Solutions", font=("Helvetica", 10),
                 bg=NAVY_HEX, fg="#9fc1f5").pack(side="left")
        tk.Label(banner, text="www.nexker.com   |   nexkerwebsolutions@gmail.com   |   00 91 8281829736",
                 font=("Helvetica", 9), bg=NAVY_HEX, fg="#cfe0ff").pack(side="right", padx=12)

        body = tk.Frame(wrap, bg="#eef1f6")
        body.pack(fill="both", expand=True, padx=10, pady=8)

        left = tk.Frame(body, bg="#eef1f6", width=440)
        left.pack(side="left", fill="y", padx=(0, 8))
        left.pack_propagate(False)

        right = tk.Frame(body, bg="#eef1f6")
        right.pack(side="left", fill="both", expand=True)

        # ---------- Invoice details ----------
        f1 = ttk.LabelFrame(left, text=" Invoice Details ")
        f1.pack(fill="x", pady=(0, 8))
        f1.columnconfigure(1, weight=1)
        f1.columnconfigure(3, weight=1)

        self.v_inv = tk.StringVar()
        self.v_date = tk.StringVar()
        self.v_due = tk.StringVar()
        self.v_cur = tk.StringVar(value="INR")
        self.v_tax = tk.StringVar(value="18")
        self.v_disc = tk.StringVar(value="0")
        self.v_adv = tk.StringVar(value="0")

        def row(parent, r, label, var, widget=None, col=0):
            ttk.Label(parent, text=label).grid(row=r, column=col, sticky="e", padx=(8, 4), pady=3)
            if widget is None:
                widget = ttk.Entry(parent, textvariable=var, width=16)
            widget.grid(row=r, column=col + 1, sticky="we", padx=(0, 10), pady=3)

        row(f1, 0, "Invoice No", self.v_inv)
        row(f1, 0, "Invoice Date", self.v_date, col=2)
        row(f1, 1, "Due Date", self.v_due)
        row(f1, 1, "Currency", self.v_cur,
            widget=ttk.Combobox(f1, textvariable=self.v_cur, values=list(CURRENCIES),
                                state="readonly", width=14), col=2)
        row(f1, 2, "Tax % (GST/VAT)", self.v_tax,
            widget=ttk.Spinbox(f1, textvariable=self.v_tax, from_=0, to=100,
                               increment=0.5, width=14))
        row(f1, 2, "Discount (amount)", self.v_disc, col=2)
        row(f1, 3, "Advance Paid (amount)", self.v_adv)
        self.lbl_taxname = ttk.Label(f1, text="", foreground="#666")
        self.lbl_taxname.grid(row=3, column=2, columnspan=2, sticky="w", padx=(4, 8))

        self.v_cur.trace_add("write", lambda *_: self._apply_currency())
        for v in (self.v_tax, self.v_disc, self.v_adv):
            v.trace_add("write", self._recalc)

        # ---------- Client ----------
        f2 = ttk.LabelFrame(left, text=" Client Details ")
        f2.pack(fill="x", pady=(0, 8))
        f2.columnconfigure(1, weight=1)

        self.v_cname = tk.StringVar()
        self.v_ccomp = tk.StringVar()
        self.v_cphone = tk.StringVar()
        self.v_cemail = tk.StringVar()
        for r, (lbl, var) in enumerate([("Client Name *", self.v_cname),
                                        ("Company", self.v_ccomp),
                                        ("Phone (WhatsApp)", self.v_cphone),
                                        ("Email", self.v_cemail)]):
            ttk.Label(f2, text=lbl).grid(row=r, column=0, sticky="e", padx=(8, 4), pady=3)
            ttk.Entry(f2, textvariable=var).grid(row=r, column=1, sticky="we", padx=(0, 10), pady=3)
        ttk.Label(f2, text="Address").grid(row=4, column=0, sticky="ne", padx=(8, 4), pady=3)
        self.t_addr = tk.Text(f2, height=3, font=("Helvetica", 9), relief="solid", bd=1)
        self.t_addr.grid(row=4, column=1, sticky="we", padx=(0, 10), pady=(3, 8))

        # ---------- Payment ----------
        f3 = ttk.LabelFrame(left, text=" Payment Details (printed on invoice) ")
        f3.pack(fill="both", expand=True, pady=(0, 2))
        f3.columnconfigure(1, weight=1)
        self.bank_vars = {}
        for r, key in enumerate(["Bank Name", "A/C Holder", "A/C Number", "IFSC / SWIFT", "UPI ID"]):
            v = tk.StringVar()
            self.bank_vars[key] = v
            ttk.Label(f3, text=key).grid(row=r, column=0, sticky="e", padx=(8, 4), pady=3)
            ttk.Entry(f3, textvariable=v).grid(row=r, column=1, sticky="we", padx=(0, 10), pady=3)
        ttk.Label(f3, text="(Only filled fields are printed)", foreground="#888")\
            .grid(row=5, column=1, sticky="w", pady=(0, 6))

        # ---------- Items ----------
        f4 = ttk.LabelFrame(right, text=" Services / Items ")
        f4.pack(fill="both", expand=True, pady=(0, 8))
        f4.columnconfigure(0, weight=1)

        self.v_service = tk.StringVar()
        cb = ttk.Combobox(f4, textvariable=self.v_service, values=ALL_SERVICES, state="readonly")
        cb.current(0)
        cb.grid(row=0, column=0, columnspan=4, sticky="we", padx=8, pady=(8, 3))

        self.v_desc = tk.StringVar()
        self.v_qty = tk.StringVar(value="1")
        self.v_rate = tk.StringVar()
        ttk.Label(f4, text="Description").grid(row=1, column=0, sticky="w", padx=8)
        ttk.Label(f4, text="Qty").grid(row=1, column=1, sticky="w")
        ttk.Label(f4, text="Rate").grid(row=1, column=2, sticky="w")
        ttk.Entry(f4, textvariable=self.v_desc).grid(row=2, column=0, sticky="we", padx=(8, 6))
        ttk.Entry(f4, textvariable=self.v_qty, width=6).grid(row=2, column=1, sticky="w")
        ttk.Entry(f4, textvariable=self.v_rate, width=11).grid(row=2, column=2, sticky="w", padx=(0, 4))
        ttk.Button(f4, text="+ Add Item", command=self._add_item).grid(row=2, column=3,
                                                                       sticky="e", padx=8)
        self.root.bind("<Return>", lambda e: self._add_item() if self.root.focus_get() in
                       (f4.children.get("!entry2"), f4.children.get("!entry3")) else None)

        cols = ("service", "desc", "qty", "rate", "amount")
        self.tree = ttk.Treeview(f4, columns=cols, show="headings", height=7)
        for cid, txt, wdt, anc in [("service", "Service", 235, "w"),
                                   ("desc", "Description", 250, "w"),
                                   ("qty", "Qty", 45, "center"),
                                   ("rate", "Rate", 95, "e"),
                                   ("amount", "Amount", 105, "e")]:
            self.tree.heading(cid, text=txt)
            self.tree.column(cid, width=wdt, anchor=anc)
        self.tree.grid(row=3, column=0, columnspan=4, sticky="nsew", padx=8, pady=(8, 4))
        vsb = ttk.Scrollbar(f4, orient="vertical", command=self.tree.yview)
        vsb.grid(row=3, column=4, sticky="ns", pady=(8, 4))
        self.tree.configure(yscrollcommand=vsb.set)
        f4.rowconfigure(3, weight=1)

        bot = ttk.Frame(f4)
        bot.grid(row=4, column=0, columnspan=5, sticky="we", padx=8, pady=(0, 8))
        ttk.Button(bot, text="Remove Selected", command=self._remove_item).pack(side="left")
        ttk.Button(bot, text="Clear All", command=self._clear_items).pack(side="left", padx=6)

        tot = ttk.Frame(bot)
        tot.pack(side="right")
        self.lbl_sub = ttk.Label(tot, text="")
        self.lbl_disc = ttk.Label(tot, text="")
        self.lbl_tax = ttk.Label(tot, text="")
        self.lbl_total = ttk.Label(tot, text="", font=("Helvetica", 10, "bold"), foreground=NAVY_HEX)
        for r, (lbl, w) in enumerate([("Subtotal:", self.lbl_sub), ("Discount:", self.lbl_disc),
                                      ("Tax:", self.lbl_tax), ("TOTAL:", self.lbl_total)]):
            ttk.Label(tot, text=lbl, font=("Helvetica", 9, "bold") if r == 3 else ())\
                .grid(row=r, column=0, sticky="e", padx=4, pady=1)
            w.grid(row=r, column=1, sticky="w", pady=1)

        # ---------- Notes & Terms ----------
        f5 = ttk.LabelFrame(right, text=" Notes & Terms ")
        f5.pack(fill="x", pady=(0, 4))
        f5.columnconfigure(1, weight=1)
        ttk.Label(f5, text="Notes").grid(row=0, column=0, sticky="nw", padx=(8, 4), pady=4)
        self.t_notes = tk.Text(f5, height=2, font=("Helvetica", 9), relief="solid", bd=1)
        self.t_notes.grid(row=0, column=1, sticky="we", padx=(0, 8), pady=4)
        ttk.Label(f5, text="Terms").grid(row=1, column=0, sticky="nw", padx=(8, 4), pady=(0, 6))
        self.t_terms = tk.Text(f5, height=4, font=("Helvetica", 9), relief="solid", bd=1)
        self.t_terms.grid(row=1, column=1, sticky="we", padx=(0, 8), pady=(0, 8))

        # ---------- Action buttons ----------
        actions = tk.Frame(wrap, bg="#eef1f6")
        actions.pack(fill="x", padx=10, pady=(0, 4))

        def abtn(text, cmd, primary=False):
            tk.Button(actions, text=text, command=cmd, cursor="hand2",
                      font=("Helvetica", 9, "bold" if primary else "normal"),
                      bg=NAVY_HEX if primary else "white",
                      fg="white" if primary else NAVY_HEX,
                      activebackground="#1b3a7a" if primary else "#e8edf5",
                      activeforeground="white" if primary else NAVY_HEX,
                      relief="flat" if primary else "groove", bd=1, padx=12, pady=6)\
                .pack(side="left", padx=(0, 6))

        abtn("Generate & Save PDF  (Ctrl+G)", self.generate, primary=True)
        abtn("Print", self._print_pdf)
        abtn("Send via WhatsApp", self._send_whatsapp)
        abtn("Send via Email", self._send_email)
        abtn("Upload to Google Drive", self._upload_drive)
        abtn("Open PDF", lambda: open_file(self.pdf_path) if self._require_pdf() else None)

        self.status = tk.Label(wrap, text=" Ready.", anchor="w", bg="#dfe5ef", fg="#333",
                               font=("Helvetica", 8))
        self.status.pack(fill="x", side="bottom")

    # ------------------------- form defaults -------------------------
    def _load_form_defaults(self):
        self.v_inv.set(self._auto_inv)
        today = datetime.now()
        self.v_date.set(today.strftime("%d-%m-%Y"))
        self.v_due.set((today + timedelta(days=7)).strftime("%d-%m-%Y"))
        self.v_cur.set(self.settings.get("last_currency", "INR"))
        self.t_notes.insert("1.0", self.settings.get("default_notes", ""))
        self.t_terms.insert("1.0", self.settings.get("default_terms", ""))
        for key, var in self.bank_vars.items():
            var.set(self.settings.get("bank_" + slug(key), ""))
        drive = self.settings.get("drive_folder_id", "")
        self._set_status("Ready.   Drive folder: " + (drive if drive else "My Drive (root)"))

    def _next_inv_no(self):
        return "NXK-{}-{:04d}".format(datetime.now().year,
                                      self.settings.get("invoice_counter", 1))

    def _apply_currency(self):
        cur = self.v_cur.get()
        if cur in CURRENCIES:
            self.lbl_taxname.config(text=CURRENCIES[cur]["tax_label"] + " applies")
            self.v_tax.set(str(CURRENCIES[cur]["default_tax"]))
            self._refresh_tree()

    # ------------------------- items / totals -------------------------
    def _add_item(self):
        svc = self.v_service.get().strip()
        desc = self.v_desc.get().strip()
        if not svc:
            return
        if svc == "Custom / Other Service":
            if not desc:
                messagebox.showwarning("Service", "Type the custom service name in the Description box.")
                return
            svc, desc = desc, ""
        try:
            qty = float(self.v_qty.get() or 1)
            rate = float(self.v_rate.get() or 0)
        except ValueError:
            messagebox.showwarning("Numbers", "Qty and Rate must be numbers.")
            return
        self.items.append({"service": svc, "desc": desc, "qty": qty, "rate": rate})
        self.v_desc.set("")
        self.v_rate.set("")
        self.v_qty.set("1")
        self.v_service.set(ALL_SERVICES[0])
        self._refresh_tree()

    def _refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        cur = self.v_cur.get()
        for i, it in enumerate(self.items):
            self.tree.insert("", "end", iid=str(i), values=(
                it["service"], it["desc"], fnum(it["qty"]),
                fmt_money(it["rate"], cur), fmt_money(it["qty"] * it["rate"], cur)))
        self._recalc()

    def _remove_item(self):
        for iid in sorted(self.tree.selection(), key=lambda x: -int(x)):
            self.items.pop(int(iid))
        self._refresh_tree()

    def _clear_items(self):
        self.items.clear()
        self._refresh_tree()

    def _recalc(self, *_):
        cur = self.v_cur.get()
        sub = sum(i["qty"] * i["rate"] for i in self.items)
        try:
            disc = max(float(self.v_disc.get() or 0), 0)
        except ValueError:
            disc = 0
        try:
            tax_pct = max(float(self.v_tax.get() or 0), 0)
        except ValueError:
            tax_pct = 0
        tax = max(sub - disc, 0) * tax_pct / 100.0
        total = max(sub - disc, 0) + tax
        self.totals = {"total": total}
        self.lbl_sub.config(text=fmt_money(sub, cur))
        self.lbl_disc.config(text="- " + fmt_money(disc, cur))
        self.lbl_tax.config(text="{} ({}%)".format(fmt_money(tax, cur), fnum(tax_pct)))
        self.lbl_total.config(text=fmt_money(total, cur))

    # ------------------------- generate -------------------------
    def generate(self, event=None):
        client = self.v_cname.get().strip()
        if not client:
            messagebox.showwarning("Missing", "Please enter the client name.")
            return
        if not self.items:
            messagebox.showwarning("Missing", "Please add at least one service / item.")
            return
        try:
            disc = float(self.v_disc.get() or 0)
            adv = float(self.v_adv.get() or 0)
            tax_pct = float(self.v_tax.get() or 0)
        except ValueError:
            messagebox.showwarning("Numbers", "Discount / Advance / Tax must be numbers.")
            return

        inv_no = self.v_inv.get().strip() or self._auto_inv
        folder = self.settings.get("save_folder") or os.path.join(
            os.path.expanduser("~"), "Documents", "NEXKER Invoices")
        os.makedirs(folder, exist_ok=True)
        path = os.path.join(folder, "Invoice_{}_{}.pdf".format(
            inv_no.replace("/", "-"), sanitize_filename(client)))

        data = {
            "inv_no": inv_no,
            "date_str": self.v_date.get().strip() or datetime.now().strftime("%d-%m-%Y"),
            "due_str": self.v_due.get().strip(),
            "currency": self.v_cur.get(),
            "client_name": client,
            "client_company": self.v_ccomp.get().strip(),
            "client_phone": self.v_cphone.get().strip(),
            "client_email": self.v_cemail.get().strip(),
            "client_address": self.t_addr.get("1.0", "end").strip(),
            "items": [dict(i) for i in self.items],
            "disc": disc, "adv": adv, "tax_pct": tax_pct,
            "notes": self.t_notes.get("1.0", "end").strip(),
            "terms": self.t_terms.get("1.0", "end").strip(),
            "bank": {k: v.get().strip() for k, v in self.bank_vars.items()},
            "tax_id": ("GSTIN: " + self.settings["gstin"]) if self.settings.get("gstin") else "",
        }
        try:
            build_invoice_pdf(data, path)
        except Exception as e:
            messagebox.showerror("PDF Error", "Could not create PDF:\n{}".format(e))
            return

        self.pdf_path = path
        if inv_no == self._auto_inv:
            self.settings["invoice_counter"] = self.settings.get("invoice_counter", 1) + 1
            self._auto_inv = self._next_inv_no()
        save_settings(self.settings)
        self._set_status("Saved: " + path)
        if messagebox.askyesno("Invoice Saved", "Invoice saved to:\n{}\n\nOpen the PDF now?".format(path)):
            open_file(path)

    def _new_invoice(self):
        for v in (self.v_cname, self.v_ccomp, self.v_cphone, self.v_cemail):
            v.set("")
        self.t_addr.delete("1.0", "end")
        self._clear_items()
        self.v_disc.set("0")
        self.v_adv.set("0")
        self.v_inv.set(self._auto_inv)
        today = datetime.now()
        self.v_date.set(today.strftime("%d-%m-%Y"))
        self.v_due.set((today + timedelta(days=7)).strftime("%d-%m-%Y"))
        self._set_status("New invoice started.")

    # ------------------------- actions -------------------------
    def _require_pdf(self):
        if self.pdf_path and os.path.exists(self.pdf_path):
            return True
        messagebox.showwarning("No invoice", "Generate & save the PDF first.")
        return False

    def _set_status(self, text):
        self.status.config(text=" " + text)

    def _print_pdf(self):
        if not self._require_pdf():
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(self.pdf_path, "print")
            elif sys.platform == "darwin":
                subprocess.Popen(["lp", self.pdf_path])
            else:
                subprocess.Popen(["lpr", self.pdf_path])
            self._set_status("Sent to printer.")
        except Exception:
            open_file(self.pdf_path)
            messagebox.showinfo("Print", "Could not reach the printer directly.\n"
                                         "The PDF is opened - press Ctrl+P in your PDF viewer.")

    def _compose_message(self):
        cur = self.v_cur.get()
        return "\n".join([
            "Hello {},".format(self.v_cname.get().strip() or "there"),
            "",
            "Thank you for choosing NEXKER Web & Digital Solutions.",
            "Please find your invoice attached (PDF).",
            "",
            "Invoice No : {}".format(self.v_inv.get().strip()),
            "Total      : {}".format(fmt_money(self.totals.get("total", 0), cur)),
            "Due Date   : {}".format(self.v_due.get().strip()),
            "",
            "Payment details are inside the invoice.",
            "",
            "Website   : www.nexker.com",
            "Instagram : instagram.com/nexke_r",
            "Facebook  : facebook.com/nexkerwebsolutions",
            "",
            "- NEXKER Web & Digital Solutions",
        ])

    def _send_whatsapp(self):
        if not self._require_pdf():
            return
        phone = re.sub(r"\D", "", self.v_cphone.get())
        if not phone:
            phone = simpledialog.askstring(
                "WhatsApp", "Client WhatsApp number (with country code,\ne.g. 919876543210):",
                parent=self.root)
            if not phone:
                return
            phone = re.sub(r"\D", "", phone)
        phone = phone.lstrip("0")
        if len(phone) == 10:
            phone = "91" + phone
        url = ("https://web.whatsapp.com/send?phone=" + phone +
               "&text=" + urllib.parse.quote(self._compose_message()))
        webbrowser.open(url)
        reveal_file(self.pdf_path)
        messagebox.showinfo(
            "WhatsApp",
            "WhatsApp Web is opening with your message pre-typed.\n\n"
            "1. Press Enter to send the message.\n"
            "2. The invoice folder is open with the PDF selected -\n"
            "   drag it into the chat and press Enter.\n\n"
            "(WhatsApp Web must be linked with your phone.)")

    def _send_email(self):
        if not self._require_pdf():
            return
        to = self.v_cemail.get().strip()
        if not to:
            to = simpledialog.askstring("Email", "Recipient email:", parent=self.root)
            if not to:
                return
        sender = BUSINESS["email"]
        pw = self.settings.get("email_app_password", "")
        if not pw:
            pw = simpledialog.askstring(
                "Gmail App Password",
                "Enter your Gmail App Password (16 characters).\n"
                "Create one at: myaccount.google.com/apppasswords",
                show="*", parent=self.root)
            if not pw:
                return
            if messagebox.askyesno("Remember?", "Save this App Password on this computer?\n"
                                   "(Stored in nexker_settings.json)"):
                self.settings["email_app_password"] = pw
                save_settings(self.settings)

        path, subject = self.pdf_path, "Invoice {} - NEXKER Web & Digital Solutions".format(
            self.v_inv.get().strip())
        body = self._compose_message() + "\n\n(The invoice PDF is attached.)"

        def work():
            try:
                msg = MIMEMultipart()
                msg["From"], msg["To"], msg["Subject"] = sender, to, subject
                msg.attach(MIMEText(body, "plain"))
                with open(path, "rb") as f:
                    att = MIMEApplication(f.read(), _subtype="pdf")
                    att.add_header("Content-Disposition", "attachment",
                                   filename=os.path.basename(path))
                    msg.attach(att)
                with smtplib.SMTP("smtp.gmail.com", 587, timeout=40) as s:
                    s.starttls()
                    s.login(sender, pw)
                    s.send_message(msg)
                self.root.after(0, lambda: self._set_status("Email sent to " + to))
                self.root.after(0, lambda: messagebox.showinfo("Email", "Invoice emailed to " + to))
            except smtplib.SMTPAuthenticationError:
                self.root.after(0, lambda: messagebox.showerror(
                    "Email", "Login failed.\n\nUse your GMAIL APP PASSWORD (16 characters),\n"
                             "not your normal password. See Help > Setup Guide."))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Email", "Send failed:\n{}".format(e)))
        threading.Thread(target=work, daemon=True).start()

    # ------------------------- drive -------------------------
    def _upload_drive(self):
        if not self._require_pdf():
            return
        self._set_status("Uploading to Google Drive ...")
        path = self.pdf_path
        folder_id = self.settings.get("drive_folder_id", "")

        def work():
            try:
                link = upload_to_gdrive(path, folder_id)
                self.root.after(0, lambda: self._set_status("Uploaded to Google Drive"))
                self.root.after(0, lambda: messagebox.showinfo(
                    "Google Drive", "Uploaded successfully!\n\nLink:\n" + link))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Google Drive", str(e)))
        threading.Thread(target=work, daemon=True).start()

    def _pick_drive_folder(self):
        self._set_status("Connecting to Google Drive ...")

        def work():
            try:
                svc = _drive_service()
                res = svc.files().list(
                    q="mimeType='application/vnd.google-apps.folder' and trashed=false",
                    spaces="drive", fields="files(id, name)", pageSize=60,
                    orderBy="name").execute()
                folders = res.get("files", [])
                self.root.after(0, lambda: self._folder_dialog(folders))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Google Drive", str(e)))
        threading.Thread(target=work, daemon=True).start()

    def _folder_dialog(self, folders):
        win = tk.Toplevel(self.root)
        win.title("Choose Google Drive Folder")
        win.geometry("440x440")
        ttk.Label(win, text="Invoices will be uploaded to the selected folder:")\
            .pack(anchor="w", padx=10, pady=(10, 4))
        lb = tk.Listbox(win, font=("Helvetica", 10))
        lb.pack(fill="both", expand=True, padx=10)
        lb.insert("end", "My Drive  (root)")
        for f in folders:
            lb.insert("end", f["name"])

        def ok(_e=None):
            sel = lb.curselection()
            if not sel:
                return
            if sel[0] == 0:
                self.settings["drive_folder_id"] = ""
                name = "My Drive (root)"
            else:
                self.settings["drive_folder_id"] = folders[sel[0] - 1]["id"]
                name = folders[sel[0] - 1]["name"]
            save_settings(self.settings)
            self._set_status("Drive folder: " + name)
            win.destroy()

        lb.bind("<Double-Button-1>", ok)
        ttk.Button(win, text="Select", command=ok).pack(pady=8)

    def _clear_token(self):
        if os.path.exists("token.json"):
            if messagebox.askyesno("Google Drive", "Delete saved Drive login (token.json)?"):
                os.remove(TOKEN_FILE)
                self._set_status("Drive login cleared.")
        else:
            messagebox.showinfo("Google Drive", "No saved login found.")

    # ------------------------- misc -------------------------
    def _choose_folder(self):
        d = filedialog.askdirectory(title="Choose where invoices are saved",
                                    initialdir=self.settings.get("save_folder") or os.path.expanduser("~"))
        if d:
            self.settings["save_folder"] = d
            save_settings(self.settings)
            self._set_status("Save folder: " + d)

    def _show_guide(self):
        win = tk.Toplevel(self.root)
        win.title("Setup Guide")
        win.geometry("640x560")
        txt = tk.Text(win, wrap="word", font=("Courier", 9), padx=10, pady=10)
        txt.pack(fill="both", expand=True, side="left")
        sb = ttk.Scrollbar(win, command=txt.yview)
        sb.pack(fill="y", side="right")
        txt.configure(yscrollcommand=sb.set)
        txt.insert("1.0", GUIDE)
        txt.config(state="disabled")

    def _about(self):
        messagebox.showinfo(
            "About", "NEXKER Invoice Generator v1.0\n\nNEXKER Web & Digital Solutions\n"
                     "www.nexker.com\nnexkerwebsolutions@gmail.com\n00 91 8281829736\n\n"
                     "Instagram: instagram.com/nexke_r\nFacebook: facebook.com/nexkerwebsolutions")

    def _on_close(self):
        for key, var in self.bank_vars.items():
            self.settings["bank_" + slug(key)] = var.get().strip()
        self.settings["last_currency"] = self.v_cur.get()
        self.settings["default_notes"] = self.t_notes.get("1.0", "end").strip()
        self.settings["default_terms"] = self.t_terms.get("1.0", "end").strip()
        save_settings(self.settings)
        self.root.destroy()


GUIDE = """NEXKER INVOICE GENERATOR - SETUP GUIDE
=========================================

1) GOOGLE DRIVE BACKUP (one-time setup)
----------------------------------------
 a. Install:  pip install google-api-python-client google-auth-oauthlib
 b. Go to console.cloud.google.com  ->  create a project (any name).
 c. "APIs & Services" -> "Library" -> search "Google Drive API" -> Enable.
 d. "OAuth consent screen" -> External -> fill in App name ->
    add your Gmail under "Test users" -> Save.
 e. "Credentials" -> "Create credentials" -> "OAuth client ID" ->
    Application type: Desktop app -> Create -> Download JSON.
 f. Rename the file to  credentials.json  and keep it in the same
    folder as this program.
 g. In the app:  Google Drive > Choose Google Drive Folder...
    (The first time, a browser window opens - sign in and click Allow.
    A token.json is saved, so you won't be asked again.)

2) EMAIL (Gmail)
-----------------
 a. Turn on 2-Step Verification for your Google account.
 b. Visit  myaccount.google.com/apppasswords
 c. Create an App password (16 characters).
 d. Paste it in the app when sending your first email.

3) WHATSAPP
------------
 - Opens WhatsApp Web with your message pre-typed, and opens the
   folder with the invoice PDF selected.
 - Press Enter to send the text, then drag the PDF into the chat.
 - Client number must include country code (e.g. 919876543210).

4) PRINT
---------
 - Windows: prints with your default printer / PDF app.
 - Mac/Linux: uses the standard lp / lpr command.

Files created next to the program:
 - nexker_settings.json   your saved preferences
 - credentials.json       (you add this) Google Drive access
 - token.json             created automatically after Drive login
"""


def main():
    root = tk.Tk()
    InvoiceApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

