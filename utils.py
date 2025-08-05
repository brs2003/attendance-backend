import os
import json
import re
from datetime import datetime

LOG_PATH = "backend/data/log.json"

def get_date_str(path):
    name = os.path.basename(path)

    match = re.search(r"(\d{2})[-_](\d{2})[-_](\d{4})", name)
    if match:
        try:
            return datetime.strptime(match.group(0), "%d-%m-%Y")
        except:
            pass

    match = re.search(r"(\d{4})[-_](\d{2})[-_](\d{2})", name)
    if match:
        try:
            return datetime.strptime(match.group(0), "%Y-%m-%d")
        except:
            pass

    print("⚠️ Could not extract date from filename, using today.")
    return datetime.today()

def is_weekend(date):
    return date.weekday() >= 5

def get_month_folder(date):
    return os.path.join("backend/data", date.strftime("%Y-%m"))

def load_log():
    if os.path.exists(LOG_PATH):
        with open(LOG_PATH, "r") as f:
            return json.load(f)
    return {}

def save_log(log):
    os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)
    with open(LOG_PATH, "w") as f:
        json.dump(log, f, indent=2)
