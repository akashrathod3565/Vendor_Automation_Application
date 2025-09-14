import os
import json
import csv
import time
import threading
import re
import traceback
import subprocess
import shutil
from datetime import datetime, timedelta
from collections import defaultdict
from typing import Optional, List, Dict, Any

import tkinter as tk
from tkinter import filedialog, messagebox
import webbrowser

import pandas as pd
from apscheduler.schedulers.background import BackgroundScheduler
import win32com.client
from win32com.client import constants
import pythoncom

from tkhtmlview import HTMLLabel

# ---------------- CONFIG ----------------
SUPPLIERS_BASE_DIR = os.getenv("SUPPLIERS_BASE_DIR", "./Suppliers")
LOG_CSV = os.path.join(SUPPLIERS_BASE_DIR, "activity_log.csv")
LAST_FETCH_FILE = os.path.join(SUPPLIERS_BASE_DIR, "last_fetch.json")
AUTO_FETCH_TIMES = ["09:00", "18:00"]

# Make sure base exists
os.makedirs(SUPPLIERS_BASE_DIR, exist_ok=True)

# ---------------- Utilities ----------------
_ILLEGAL_FN_RE = re.compile(r'[^A-Za-z0-9_.-]')
_EMAIL_RE = re.compile(r'[\w\.-]+@[\w\.-]+')

def sanitize_filename(s: str, maxlen: int = 120) -> str:
    if not s:
        return "unknown"
    s = re.sub(r"\s+", "_", s.strip())
    s = _ILLEGAL_FN_RE.sub("_", s)
    if len(s) > maxlen:
        base, dot, ext = s.rpartition(".")
        if dot and ext:
            base = base[: maxlen - len(ext) - 1]
            s = f"{base}.{ext}"
        else:
            s = s[:maxlen]
    return s

def safe_local_path_for_msg(outbox_dir: str, prefix: str = "rfq_") -> str:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    fname = f"{prefix}{ts}.msg"
    fname = sanitize_filename(fname, maxlen=120)
    path = os.path.join(outbox_dir, fname)
    # keep path length reasonable
    if len(path) > 240:
        fname = fname[:80] + ".msg"
        path = os.path.join(outbox_dir, sanitize_filename(fname))
    return path

def ensure_supplier_dirs_unique(supplier_name: str, vendor_email: str):
    local = vendor_email.split("@")[0] if "@" in vendor_email else vendor_email
    if supplier_name:
        folder = f"{supplier_name}_{local}"
    else:
        folder = local
    folder = sanitize_filename(folder, maxlen=120)
    base = os.path.join(SUPPLIERS_BASE_DIR, folder)
    outbox = os.path.join(base, "Outbox")
    quotes = os.path.join(base, "Quotations")
    os.makedirs(outbox, exist_ok=True)
    os.makedirs(quotes, exist_ok=True)
    return base, outbox, quotes

def append_log(action: str, supplier: str, vendor_emails: List[str], details: str = ""):
    header = ["timestamp", "action", "supplier", "vendor_emails", "details"]
    row = [datetime.now().isoformat(), action, supplier, ";".join(vendor_emails), details]
    write_header = not os.path.exists(LOG_CSV)
    try:
        with open(LOG_CSV, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            if write_header:
                writer.writerow(header)
            writer.writerow(row)
    except Exception:
        print("Failed to write log:", traceback.format_exc())

def load_last_fetch():
    if os.path.exists(LAST_FETCH_FILE):
        try:
            with open(LAST_FETCH_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    default = (datetime.now() - timedelta(days=7)).isoformat()
    return {"last_fetch": default}

def save_last_fetch(ts_iso: str):
    try:
        with open(LAST_FETCH_FILE, "w", encoding="utf-8") as f:
            json.dump({"last_fetch": ts_iso}, f)
    except Exception:
        print("Failed to save last_fetch:", traceback.format_exc())

def format_for_outlook_restrict(dt: datetime):
    return dt.strftime("%m/%d/%Y %I:%M %p")

def extract_smtp_from_sender(mail) -> Optional[str]:
    """
    Robust extraction of SMTP/primary address from various Outlook Sender representations.
    Returns lowercased email string or None.
    """
    try:
        raw = ""
        try:
            raw = getattr(mail, "SenderEmailAddress", "") or ""
        except Exception:
            raw = ""
        raw = str(raw).strip()
        m = _EMAIL_RE.search(raw)
        if m:
            return m.group(0).lower()

        sender_obj = None
        try:
            sender_obj = getattr(mail, "Sender", None)
        except Exception:
            sender_obj = None

        if sender_obj:
            try:
                exch = sender_obj.GetExchangeUser()
                if exch:
                    addr = exch.PrimarySmtpAddress
                    if addr:
                        return str(addr).lower()
            except Exception:
                pass

            try:
                addr = getattr(sender_obj, "Address", "") or ""
                m = _EMAIL_RE.search(str(addr))
                if m:
                    return m.group(0).lower()
            except Exception:
                pass

        try:
            sname = getattr(mail, "SenderName", "") or ""
            m = _EMAIL_RE.search(str(sname))
            if m:
                return m.group(0).lower()
        except Exception:
            pass

    except Exception:
        pass
    return None

# ---------------- Excel loader ----------------
def load_vendors_from_excel(path: str) -> Dict[str, List[Dict[str, Any]]]:
    df = pd.read_excel(path, engine="openpyxl")
    cols = {c.lower().strip(): c for c in df.columns}

    supplier_col = cols.get("suppliername") or cols.get("supplier")
    vendor_email_col = cols.get("vendoremail") or cols.get("email") or cols.get("vendor")
    vendor_name_col = cols.get("vendorname")
    vendor_addr_col = cols.get("vendoraddress") or cols.get("address")
    cc_col = cols.get("cc")

    if supplier_col is None or vendor_email_col is None:
        raise ValueError("Excel must contain SupplierName and VendorEmail columns (case-insensitive).")

    vendor_map: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for _, row in df.iterrows():
        sup = str(row[supplier_col]).strip()
        ven_raw = str(row[vendor_email_col]).strip()
        ven_name = str(row[vendor_name_col]).strip() if vendor_name_col else ""
        ven_addr = str(row[vendor_addr_col]).strip() if vendor_addr_col else ""
        ven_cc = str(row[cc_col]).strip() if cc_col else ""

        if sup and ven_raw and ven_raw.lower() != "nan":
            for e in [x.strip().lower() for x in ven_raw.split(",") if x.strip()]:
                vendor_map[sup].append({
                    "email": e,
                    "name": ven_name,
                    "address": ven_addr,
                    "cc": ven_cc
                })
    return vendor_map

# ---------------- Outlook fetch ----------------
def fetch_emails_from_outlook(vendor_map: dict, manual_email: Optional[str], ui_log_callback):
    def worker():
        pythoncom.CoInitialize()
        try:
            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
                ns = outlook.GetNamespace("MAPI")
                inbox = ns.GetDefaultFolder(6)
            except Exception:
                ui_log_callback("Could not connect to Outlook. Make sure Outlook is running.")
                return

            last_fetch_info = load_last_fetch()
            try:
                last_fetch_time = datetime.fromisoformat(last_fetch_info["last_fetch"])
            except Exception:
                last_fetch_time = datetime.now() - timedelta(days=7)

            restriction = "[ReceivedTime] >= '{}'".format(format_for_outlook_restrict(last_fetch_time))

            try:
                items = inbox.Items.Restrict(restriction)
            except Exception:
                try:
                    items = inbox.Items
                except Exception:
                    ui_log_callback("Could not access inbox items.")
                    return

            try:
                items.Sort("[ReceivedTime]", True)
            except Exception:
                pass

            if manual_email:
                emails_to_check = {manual_email.lower(): ("Manual", {"email": manual_email.lower()})}
            else:
                emails_to_check = {v["email"].lower(): (s, v) for s, vendors in vendor_map.items() for v in vendors}

            # copy items since Outlook Items can be non-iterable
            try:
                item_list = [items[i] for i in range(1, items.Count + 1)]
            except Exception:
                try:
                    item_list = list(items)
                except Exception:
                    ui_log_callback("No items iterable in inbox.")
                    item_list = []

            for mail in item_list:
                try:
                    sender_smtp = extract_smtp_from_sender(mail)
                    if not sender_smtp:
                        ui_log_callback(f"Skipped a mail because sender SMTP could not be determined (subject={getattr(mail,'Subject','')}).")
                        append_log("fetch_skipped", "", [getattr(mail, 'SenderName', '')], "No SMTP")
                        continue

                    sender_smtp = sender_smtp.lower()

                    if sender_smtp in emails_to_check:
                        supplier, vendor = emails_to_check[sender_smtp]
                        base, _, quotes = ensure_supplier_dirs_unique(supplier, vendor["email"])

                        # Save message (.msg) - best-effort
                        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                        msg_path = os.path.join(base, f"mail_{ts}.msg")
                        try:
                            os.makedirs(os.path.dirname(msg_path), exist_ok=True)
                            if not msg_path.lower().endswith('.msg'):
                                msg_path += '.msg'
                            try:
                                mail.SaveAs(msg_path, constants.olMSG)
                            except Exception:
                                try:
                                    mail.SaveAs(msg_path, 3)
                                except Exception:
                                    ui_log_callback(f"Warning: Could not save .msg for {sender_smtp} — {traceback.format_exc()}")
                        except Exception:
                            ui_log_callback(f"Warning: Could not prepare path to save mail for {sender_smtp} — {traceback.format_exc()}")

                        # Save attachments robustly into Quotations folder
                        try:
                            attachments = getattr(mail, "Attachments", None)
                            if attachments:
                                saved_any = False
                                # COM collection access via index
                                try:
                                    for i in range(1, attachments.Count + 1):
                                        att = attachments.Item(i)
                                        fn = getattr(att, "FileName", None) or 'attachment'
                                        att_path = os.path.join(quotes, sanitize_filename(fn))
                                        att.SaveAsFile(att_path)
                                        ui_log_callback(f"Saved attachment for {sender_smtp}: {att_path}")
                                        saved_any = True
                                except Exception:
                                    # fallback iteration
                                    try:
                                        for att in attachments:
                                            try:
                                                fn = getattr(att, "FileName", None) or 'attachment'
                                                att_path = os.path.join(quotes, sanitize_filename(fn))
                                                att.SaveAsFile(att_path)
                                                ui_log_callback(f"Saved attachment for {sender_smtp}: {att_path}")
                                                saved_any = True
                                            except Exception:
                                                ui_log_callback(f"Failed to save one attachment for {sender_smtp}.")
                                    except Exception:
                                        ui_log_callback("Attachment save attempted but failed for some attachments.")
                                if not saved_any:
                                    ui_log_callback("Attachment save attempted but none saved.")
                        except Exception as e:
                            ui_log_callback(f"Attachment save failed: {e}")

                        ui_log_callback(f"Fetched mail from {sender_smtp} for supplier {supplier}.")
                        append_log("fetch", supplier, [sender_smtp], f"subject={getattr(mail, 'Subject', '')}")
                    else:
                        ui_log_callback(f"Skipped mail from {sender_smtp} (no match).")
                        append_log("fetch_skipped", "", [sender_smtp], f"subject={getattr(mail, 'Subject', '')}")
                except Exception:
                    ui_log_callback(f"Error processing mail: {traceback.format_exc()}")

            save_last_fetch(datetime.now().isoformat())
            ui_log_callback("Fetch completed.")
        finally:
            pythoncom.CoUninitialize()

    threading.Thread(target=worker, daemon=True).start()

# ---------------- Outlook send worker ----------------
def _send_rfqs_worker(vendor_map: dict, subject: str, body_template: str, attachment_path: Optional[str],
                      manual_cc: str, ui_log_callback, manual_email: Optional[str] = None, auto_send: bool = False):
    def worker():
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")

            if manual_email:
                to_list = [{"email": manual_email.strip(), "name": "", "address": "", "cc": "", "supplier": "Manual"}]
            else:
                to_list = [{"supplier": s, **v} for s, vendors in vendor_map.items() for v in vendors]

            for vendor in to_list:
                vendor_email = vendor.get("email", "").strip()
                supplier_name = vendor.get("supplier", "")

                if not vendor_email or "@" not in vendor_email:
                    ui_log_callback(f"Skipping invalid vendor email: {vendor_email}")
                    append_log("send_skipped", supplier_name, [vendor_email], "Invalid email")
                    continue

                try:
                    vendor_name = vendor.get("name", "")
                    vendor_addr = vendor.get("address", "")
                    vendor_cc = vendor.get("cc", "").strip()

                    body_vars = defaultdict(str, {
                        "SupplierName": supplier_name,
                        "VendorName": vendor_name,
                        "VendorAddress": vendor_addr,
                        "CC": vendor_cc
                    })
                    try:
                        body = body_template.format_map(body_vars)
                    except Exception:
                        body = body_template

                    # Create mail and populate
                    try:
                        mail = outlook.CreateItem(0)
                    except Exception:
                        ui_log_callback(f"Error creating Outlook mail item for {vendor_email}: {traceback.format_exc()}")
                        append_log("send_error", supplier_name, [vendor_email], "CreateItem failed")
                        continue

                    mail.To = vendor_email
                    cc_list = [x.strip() for x in [vendor_cc, manual_cc] if x and "@" in x]
                    if cc_list:
                        mail.CC = ";".join(cc_list)
                    mail.Subject = subject
                    if any(tag in body for tag in ("<p>", "<br>", "<html", "<div", "<span", "<table")):
                        mail.HTMLBody = body
                    else:
                        mail.Body = body

                    if attachment_path:
                        abs_attach = os.path.abspath(attachment_path)
                        if os.path.exists(abs_attach):
                            try:
                                mail.Attachments.Add(abs_attach)
                            except Exception:
                                ui_log_callback(f"Warning: Could not add attachment to mail for {vendor_email}: {traceback.format_exc()}")
                        else:
                            ui_log_callback(f"Attachment not found, skipping attach: {abs_attach}")

                    # Prepare local outbox path
                    try:
                        _, outbox_path, _ = ensure_supplier_dirs_unique(supplier_name, vendor_email)
                        local_msg_path = safe_local_path_for_msg(outbox_path, prefix="rfq_")

                        # Ensure directories
                        os.makedirs(os.path.dirname(local_msg_path), exist_ok=True)

                        # Always save draft first (ensures item exists in Outlook Drafts)
                        try:
                            mail.Save()
                        except Exception:
                            ui_log_callback(f"Warning: mail.Save() failed for {vendor_email}: {traceback.format_exc()}")

                        saved_ok = False
                        # Attempt 1: preferred constant
                        try:
                            mail.SaveAs(local_msg_path, constants.olMSG)
                            saved_ok = True
                            ui_log_callback(f"Draft saved for {vendor_email} at {local_msg_path}")
                            append_log("send_draft", supplier_name, [vendor_email], f"attachment={os.path.basename(attachment_path) if attachment_path else ''}")
                        except Exception:
                            # Attempt 2: numeric constant 3 (older Outlook variations)
                            try:
                                mail.SaveAs(local_msg_path, 3)
                                saved_ok = True
                                ui_log_callback(f"Draft saved for {vendor_email} at {local_msg_path} (numeric constant)")
                                append_log("send_draft", supplier_name, [vendor_email], f"attachment={os.path.basename(attachment_path) if attachment_path else ''}")
                            except Exception:
                                # Attempt 3: fallback -> write text file and save attachments separately
                                try:
                                    fallback_txt = local_msg_path[:-4] + ".txt"
                                    with open(fallback_txt, "w", encoding="utf-8") as fh:
                                        fh.write(f"To: {vendor_email}\nSubject: {mail.Subject or ''}\n\n")
                                        body_text = getattr(mail, "HTMLBody", None) or getattr(mail, "Body", None) or ""
                                        fh.write(body_text)
                                    attach_dir = os.path.join(outbox_path, "Attachments")
                                    os.makedirs(attach_dir, exist_ok=True)
                                    for att in getattr(mail, "Attachments", []):
                                        try:
                                            fn = sanitize_filename(getattr(att, "FileName", "attachment"))
                                            att_path = os.path.join(attach_dir, fn)
                                            att.SaveAsFile(att_path)
                                        except Exception:
                                            pass
                                    ui_log_callback(f"Saved fallback TXT for {vendor_email} at {fallback_txt}")
                                    append_log("send_draft_txt", supplier_name, [vendor_email], fallback_txt)
                                    saved_ok = True
                                except Exception:
                                    ui_log_callback(f"Warning: Could not save draft for {vendor_email}: {traceback.format_exc()}")
                                    append_log("send_draft_failed", supplier_name, [vendor_email], traceback.format_exc())

                    except Exception:
                        ui_log_callback(f"Warning: Could not determine outbox path for {vendor_email}: {traceback.format_exc()}")

                    # Optionally send
                    if auto_send:
                        try:
                            mail.Send()
                            ui_log_callback(f"Sent RFQ to {vendor_email}")
                            append_log("send", supplier_name, [vendor_email], "")
                        except Exception:
                            ui_log_callback(f"Send failed for {vendor_email}: {traceback.format_exc()}")
                            append_log("send_error", supplier_name, [vendor_email], traceback.format_exc())

                    # small throttle to avoid hammering Outlook
                    time.sleep(0.25)
                except Exception:
                    ui_log_callback(f"Error preparing mail for {vendor_email}: {traceback.format_exc()}")
                    append_log("send_error", supplier_name, [vendor_email], traceback.format_exc())

        finally:
            pythoncom.CoUninitialize()

    threading.Thread(target=worker, daemon=True).start()

# ---------------- APP STATE ----------------
APP_STATE = {
    "vendor_map": {},
    "vendor_excel_path": None,
    "attachment_path": None
}

# ---------------- GUI ----------------
class VendorAppGUI:
    def __init__(self, root):
        self.root = root
        root.title("Procurement Automation App — V2.6 (fixed)")

        row = 0
        tk.Label(root, text="Vendors Excel (optional):").grid(row=row, column=0, sticky="w")
        self.excel_path_var = tk.StringVar()
        tk.Entry(root, textvariable=self.excel_path_var, width=60).grid(row=row, column=1)
        tk.Button(root, text="Load Excel", command=self.load_excel).grid(row=row, column=2)
        row += 1

        tk.Label(root, text="Attachment (RFQ):").grid(row=row, column=0, sticky="w")
        self.attach_var = tk.StringVar()
        tk.Entry(root, textvariable=self.attach_var, width=60).grid(row=row, column=1)
        tk.Button(root, text="Choose", command=self.choose_attachment).grid(row=row, column=2)
        row += 1

        tk.Label(root, text="Manual Vendor Email (single):").grid(row=row, column=0, sticky="w")
        self.manual_email_entry = tk.Entry(root, width=60)
        self.manual_email_entry.grid(row=row, column=1, columnspan=2, sticky="we")
        row += 1

        tk.Label(root, text="Manual CC:").grid(row=row, column=0, sticky="w")
        self.cc_entry = tk.Entry(root, width=60)
        self.cc_entry.grid(row=row, column=1, columnspan=2, sticky="we")
        row += 1

        tk.Label(root, text="Subject:").grid(row=row, column=0, sticky="w")
        self.subject_entry = tk.Entry(root, width=60)
        self.subject_entry.grid(row=row, column=1, columnspan=2, sticky="we")
        row += 1

        tk.Label(root, text="Body (HTML supported):").grid(row=row, column=0, sticky="nw")
        self.body_text = tk.Text(root, width=60, height=8)
        self.body_text.grid(row=row, column=1, columnspan=2, sticky="we")
        self.body_text.insert("1.0", "<p>Hello <b>{VendorName}</b>,</p><p>Please confirm your company details.</p>")
        row += 1

        tk.Button(root, text="Preview Body", command=self.preview_body).grid(row=row, column=1, sticky="we")
        row += 1

        self.body_preview = HTMLLabel(root, html="<p>Preview will appear here.</p>", width=80, height=15)
        self.body_preview.grid(row=row, column=1, columnspan=2, sticky="we")
        row += 1

        self.auto_send_var = tk.BooleanVar(value=False)
        tk.Checkbutton(root, text="Auto send", variable=self.auto_send_var).grid(row=row, column=0, columnspan=3, sticky="w")
        row += 1

        tk.Button(root, text="Send All", command=self.on_send_all).grid(row=row, column=0)
        tk.Button(root, text="Fetch Now", command=self.on_fetch_now).grid(row=row, column=1)
        tk.Button(root, text="Open Folder", command=self.on_open_folder).grid(row=row, column=2)
        row += 1

        tk.Label(root, text="Logs:").grid(row=row, column=0, sticky="nw")
        self.status = tk.Text(root, width=100, height=14)
        self.status.grid(row=row, column=1, columnspan=2)
        row += 1

        self.scheduler = BackgroundScheduler()
        for t in AUTO_FETCH_TIMES:
            hour, minute = map(int, t.split(":"))
            # schedule triggers which call scheduled_fetch_job
            self.scheduler.add_job(self.scheduled_fetch_job, 'cron', hour=hour, minute=minute)
        self.scheduler.start()
        self.log("Scheduler started.")

        root.protocol("WM_DELETE_WINDOW", self.on_close)

    def log(self, text: str):
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.status.insert("end", f"[{ts}] {text}\n")
        self.status.see("end")
        print(f"[{ts}] {text}")

    def thread_safe_log(self, text: str):
        self.root.after(0, lambda: self.log(text))

    def preview_body(self):
        body_html = self.body_text.get("1.0", "end").strip()
        self.body_preview.set_html(body_html)

    def load_excel(self):
        path = filedialog.askopenfilename(title="Select vendors Excel", filetypes=[("Excel files", "*.xlsx;*.xls")])
        if not path:
            return
        try:
            vendor_map = load_vendors_from_excel(path)
            APP_STATE["vendor_map"] = vendor_map
            APP_STATE["vendor_excel_path"] = path
            self.excel_path_var.set(path)
            self.log(f"Loaded Excel: {path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.log("Excel load error: " + str(e))

    def choose_attachment(self):
        path = filedialog.askopenfilename(title="Select RFQ attachment")
        if not path:
            return
        APP_STATE["attachment_path"] = path
        self.attach_var.set(path)
        self.log("Selected attachment: " + os.path.basename(path))

    def on_send_all(self):
        manual_email = self.manual_email_entry.get().strip()
        if not manual_email and not APP_STATE.get("vendor_map"):
            messagebox.showwarning("No vendors", "Enter manual email or load vendor Excel first.")
            return
        subject = self.subject_entry.get().strip() or "Request for Quotation"
        body = self.body_text.get("1.0", "end").strip()
        attachment = APP_STATE.get("attachment_path")
        manual_cc = self.cc_entry.get().strip()
        auto_send = bool(self.auto_send_var.get())
        self.thread_safe_log("Send worker started in background.")

        _send_rfqs_worker(APP_STATE.get("vendor_map", {}), subject, body, attachment, manual_cc,
                          self.thread_safe_log, manual_email=manual_email if manual_email else None, auto_send=auto_send)

    def on_fetch_now(self):
        manual_email = self.manual_email_entry.get().strip()
        fetch_emails_from_outlook(APP_STATE.get("vendor_map", {}), manual_email if manual_email else None, self.thread_safe_log)
        self.log("Started fetch worker.")

    def _open_path_in_explorer(self, path: str):
        try:
            path = os.path.abspath(path)
            if os.name == 'nt':
                os.startfile(path)
            else:
                if shutil.which("xdg-open"):
                    subprocess.Popen(['xdg-open', path])
                else:
                    webbrowser.open('file://' + path)
            return True
        except Exception:
            try:
                subprocess.Popen(['explorer', path])
                return True
            except Exception:
                try:
                    webbrowser.open('file://' + path)
                    return True
                except Exception:
                    return False

    def on_open_folder(self):
        email = self.manual_email_entry.get().strip().lower()
        if email:
            found = False
            for supplier, vendors in APP_STATE.get("vendor_map", {}).items():
                for v in vendors:
                    if v.get("email", "").lower() == email:
                        base, _, _ = ensure_supplier_dirs_unique(supplier, email)
                        ok = self._open_path_in_explorer(base)
                        self.log(f"Opened folder: {base}" if ok else f"Could not open folder: {base}")
                        found = True
                        break
                if found:
                    break
            if not found:
                supplier_name = "Manual"
                base, _, _ = ensure_supplier_dirs_unique(supplier_name, email)
                ok = self._open_path_in_explorer(base)
                self.log(f"Opened folder: {base}" if ok else f"Could not open folder: {base}")
            return

        if not APP_STATE.get("vendor_map"):
            messagebox.showwarning("Missing info", "Enter vendor email or load Excel.")
            return

        ok = self._open_path_in_explorer(SUPPLIERS_BASE_DIR)
        self.log(f"Opened folder: {SUPPLIERS_BASE_DIR}" if ok else f"Could not open folder: {SUPPLIERS_BASE_DIR}")

    def scheduled_fetch_job(self):
        def do_fetch():
            manual_email = self.manual_email_entry.get().strip()
            fetch_emails_from_outlook(APP_STATE.get("vendor_map", {}), manual_email if manual_email else None, self.thread_safe_log)
            self.log("Scheduled fetch triggered.")
        try:
            self.root.after(0, do_fetch)
        except Exception as e:
            fetch_emails_from_outlook(APP_STATE.get("vendor_map", {}), None, self.thread_safe_log)
            self.log(f"Scheduled fetch fallback triggered (error reading UI): {e}")

    def on_close(self):
        try:
            self.scheduler.shutdown(wait=False)
        except Exception:
            pass
        try:
            self.root.destroy()
        except Exception:
            pass

# ---------------- Run App ----------------
if __name__ == "__main__":
    root = tk.Tk()
    gui_app = VendorAppGUI(root)
    root.mainloop()
