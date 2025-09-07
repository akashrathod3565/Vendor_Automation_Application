"""
VendorApp V2.2 — safe Outlook automation (no-send default), unique supplier/vendor folders,
background threads with COM init, Excel loader, twice-daily fetch scheduler, activity log.
"""

import os
import json
import csv
import time
import threading
import re
from datetime import datetime, timedelta
from collections import defaultdict
from typing import Optional, List

import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd
from apscheduler.schedulers.background import BackgroundScheduler
import win32com.client
import pythoncom

# Configuration
SUPPLIERS_BASE_DIR = os.getenv("SUPPLIERS_BASE_DIR", "./Suppliers")
LOG_CSV = os.path.join(SUPPLIERS_BASE_DIR, "activity_log.csv")
LAST_FETCH_FILE = os.path.join(SUPPLIERS_BASE_DIR, "last_fetch.json")
AUTO_FETCH_TIMES = ["09:00", "18:00"]

os.makedirs(SUPPLIERS_BASE_DIR, exist_ok=True)

# Utilities and workers defined here (as provided) [code omitted for brevity]

# APP_STATE
APP_STATE = {
    "vendor_map": {},
    "vendor_excel_path": None,
    "attachment_path": None
}

# GUI Class and methods (as provided) [code omitted for brevity]

if __name__ == "__main__":
    root = tk.Tk()
    app = VendorAppGUI(root)
    try:
        root.mainloop()
    finally:
        try:
            app.scheduler.shutdown(wait=False)
        except Exception:
            pass
