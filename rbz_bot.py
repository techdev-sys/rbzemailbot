import os
import io
import csv
import json # <--- NEW IMPORT
import hashlib
import logging
import datetime
import win32com.client
import pandas as pd
from openpyxl import load_workbook


# ADD THIS FUNCTION TO ALLOW DYNAMIC UPDATES
def update_save_path(new_path):
    """Updates the ROOT_DIR dynamically without reloading the module."""
    global ROOT_DIR, LOG_DIR, QUARANTINE_DIR, DASHBOARD_DIR, HASH_REGISTRY
    
    ROOT_DIR = new_path
    LOG_DIR = os.path.join(ROOT_DIR, "logs")
    QUARANTINE_DIR = os.path.join(ROOT_DIR, "quarantine")
    DASHBOARD_DIR = os.path.join(ROOT_DIR, "dashboards")
    HASH_REGISTRY = os.path.join(ROOT_DIR, "registry.csv")

    # Ensure the new folders exist
    for d in [ROOT_DIR, LOG_DIR, QUARANTINE_DIR, DASHBOARD_DIR]:
        os.makedirs(d, exist_ok=True)
    
    logging.info(f"Configuration updated. Saving to: {ROOT_DIR}")

# ... rest of your code ...

# =========================================================
# 1. DYNAMIC CONFIGURATION (READS USER SETTINGS)
# =========================================================

CONFIG_FILE = "user_config.json"
DEFAULT_ROOT = r"C:\RBZ_RETURNS_SYSTEM"

def get_config():
    """Reads the user config to find the save location."""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                data = json.load(f)
                return data.get("save_path", DEFAULT_ROOT)
        except:
            pass
    return DEFAULT_ROOT

# LOAD THE PATH DYNAMICALLY
ROOT_DIR = get_config()

# DEFINE SUBFOLDERS BASED ON THE CHOSEN ROOT
LOG_DIR = os.path.join(ROOT_DIR, "logs")
QUARANTINE_DIR = os.path.join(ROOT_DIR, "quarantine")
DASHBOARD_DIR = os.path.join(ROOT_DIR, "dashboards")
HASH_REGISTRY = os.path.join(ROOT_DIR, "registry.csv")

# Create folders immediately
for d in [ROOT_DIR, LOG_DIR, QUARANTINE_DIR, DASHBOARD_DIR]:
    os.makedirs(d, exist_ok=True)
    
DRY_RUN = False

# =========================================================
# 2. LOGGING
# =========================================================

log_file = os.path.join(LOG_DIR, f"intake_{datetime.date.today()}.log")
logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)
console = logging.StreamHandler()
console.setLevel(logging.INFO)
logging.getLogger('').addHandler(console)

# =========================================================
# 3. BANK RULES
# =========================================================

BANK_RULES = [
    {"id": "FBCCROWN", "sender_domains": ["@fbc.co.zw"], "subject_includes": ["CROWN"], "names": ["CROWN"]},
    {"id": "ZB_GROUP", "sender_domains": ["@zb.co.zw"], "names": ["ZB BANK", "ZB BUILDING", "ZBBS", "ZB GROUP"]},
    {"id": "FBCBANK", "sender_domains": ["@fbc.co.zw"], "names": ["FBC BANK", "FBC HOLDINGS"]},
    {"id": "TIMEBANK", "sender_email": "ninoymatcheso@gmail.com", "names": ["TIMEBANK"]},
    {"id": "EMPOWERBANK", "sender_domains": ["@empowerbank.co.zw"], "names": ["EMPOWER"]},
    {"id": "BANCABC", "sender_domains": ["@bancabc.co.zw"], "names": ["BANCABC", "BANC ABC"]},
    {"id": "STANBIC", "sender_domains": ["@stanbic.com"], "names": ["STANBIC"]},
    {"id": "AFC", "sender_domains": ["@afcholdings.co.zw"], "names": ["AFC", "AGRICULTURAL"]},
    {"id": "SUCCESS", "sender_domains": ["@successbank.co.zw"], "names": ["SUCCESS"]},
    {"id": "GETBUCKS", "sender_domains": ["@getbucksbank.com"], "names": ["GETBUCKS"]},
    {"id": "STEWARD", "sender_domains": ["@stewardbank.co.zw"], "names": ["STEWARD", "TN CYBER"]},
    {"id": "ACL", "sender_domains": ["@africancentury.co.zw"], "names": ["AFRICAN CENTURY", "ACL"]},
    {"id": "NMB", "sender_domains": ["@nmbz.co.zw"], "names": ["NMB"]},
    {"id": "NBS", "sender_domains": ["@nbs.co.zw"], "names": ["NBS", "NATIONAL BUILDING"]},
    {"id": "ZWMB", "sender_domains": ["@womensbank.co.zw"], "names": ["WOMEN", "ZWMB"]},
    {"id": "FIRSTCAPITAL", "sender_domains": ["@firstcapitalbank.co.zw"], "names": ["FIRST CAPITAL", "FIRSTCAPITAL"]},
    {"id": "METBANK", "sender_domains": ["@metbank.co.zw"], "names": ["METBANK", "MET BANK"]},
    {"id": "MUKURU", "sender_domains": ["@mukuru.com"], "names": ["MUKURU"]},
    {"id": "INNBUCKS", "sender_domains": ["@innbucks.co.zw"], "names": ["INNBUCKS"]},
    {"id": "CBZ", "sender_domains": ["@cbz.co.zw"], "names": ["CBZ"]},
    {"id": "NEDBANK", "sender_domains": ["@nedbank.co.zw"], "names": ["NEDBANK"]},
    {"id": "ECOBANK", "sender_domains": ["@ecobank.com"], "names": ["ECOBANK"]},
    {"id": "IDBZ", "sender_domains": ["@idbz.co.zw"], "names": ["IDBZ"]},
    {"id": "CABS", "sender_domains": ["@oldmutual.co.zw"], "names": ["CABS"]},
    {"id": "POSB", "sender_domains": ["@posb.co.zw"], "names": ["POSB"]}
]

EXPECTED_RETURNS = {
    "DAILY_LOANS": ["AFC", "BANCABC", "CABS", "CBZ", "ECOBANK", "FBCBANK", "FBCCROWN", "FIRSTCAPITAL", "METBANK", "NBS", "NEDBANK", "NMB", "POSB", "STANBIC", "STEWARD", "ZBBANK", "ZBBS"],
    "BSD2_3": ["EMPOWERBANK", "BANCABC", "STANBIC", "FBCBANK", "FBCCROWN", "AFC", "SUCCESS", "TIMEBANK", "GETBUCKS", "STEWARD", "ACL", "NMB", "NBS", "ZWMB", "FIRSTCAPITAL", "METBANK", "MUKURU", "INNBUCKS", "CBZ", "NEDBANK", "ECOBANK", "IDBZ", "CABS", "ZBBANK", "ZBBS", "POSB"],
    "BSD4": ["EMPOWERBANK", "BANCABC", "STANBIC", "FBCBANK", "FBCCROWN", "AFC", "SUCCESS", "TIMEBANK", "GETBUCKS", "STEWARD", "ACL", "NMB", "NBS", "ZWMB", "FIRSTCAPITAL", "METBANK", "MUKURU", "INNBUCKS", "CBZ", "NEDBANK", "ECOBANK", "IDBZ", "CABS", "ZBBANK", "ZBBS", "POSB"]
}

def identify_bank(sender_email, subject=""):
    sender_email = sender_email.lower().strip()
    subject = subject.upper()
    if "tsungai" in sender_email or "exchangelabs" in sender_email: pass 
    for rule in BANK_RULES:
        match = False
        if "sender_email" in rule and rule["sender_email"] == sender_email: match = True
        elif "sender_domains" in rule:
            domain = sender_email.split("@")[-1]
            if any(domain.endswith(d.replace("@", "")) for d in rule["sender_domains"]): match = True
        if not match:
            if "names" in rule and any(n in subject for n in rule["names"]): match = True
        if match:
            if "subject_includes" in rule and not any(k in subject for k in rule["subject_includes"]): continue
            if rule["id"] == "ZB_GROUP": 
                return "ZBBS" if ("ZBBS" in subject or "BUILDING" in subject) else "ZBBANK"
            return rule["id"]
    return "UNKNOWN_BANK"

def refine_bank_id(current_id, filename):
    filename = filename.lower()
    if current_id in ["ZBBANK", "ZBBS", "ZB_GROUP"]:
        if "building" in filename or "society" in filename or "zbbs" in filename: return "ZBBS"
        return "ZBBANK"
    if current_id in ["FBCBANK", "FBCCROWN"]:
        if "building" in filename or "society" in filename: return "FBCBANK" 
    return current_id

def analyze_excel(file_bytes):
    try:
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
        sheet_names = [s.upper() for s in wb.sheetnames]
        sheet_count = len(sheet_names)
        ws = wb.active
        val_b1 = str(ws.cell(1, 2).value).upper()
        val_a1 = str(ws.cell(1, 1).value).upper()
        has_loans     = any("LOAN" in s for s in sheet_names)
        has_deposits  = any("DEPOSIT" in s for s in sheet_names)
        has_liquidity = any("LIQUIDITY" in s for s in sheet_names)
        if has_liquidity: return "DAILY_LOANS", None
        if has_loans and has_deposits:
            if 2 <= sheet_count <= 5: return "DAILY_LOANS", None
        if "BSD4" in val_b1 or "BSD 4" in val_b1: return "BSD4", None
        if any(k in s for s in sheet_names for k in ["BSD", "FOREIGN", "CURRENCY", "EXPOSURE"]): return "BSD2_3", None
        if "FOREIGN" in val_a1 or "CURRENCY" in val_a1 or "BSD 2" in val_a1: return "BSD2_3", None
    except: pass
    return "UNKNOWN", None

def save_payload(payload, bank_hint, rtype, use_date):
    if rtype == "UNKNOWN": return False
    file_hash = hashlib.sha256(payload).hexdigest()
    path = os.path.join(ROOT_DIR, rtype, str(use_date.year), use_date.strftime("%B"), use_date.strftime("%Y-%m-%d"))
    if DRY_RUN: return True
    os.makedirs(path, exist_ok=True)
    base_name = bank_hint
    extension = ".xlsx"
    counter = 1
    file_path = os.path.join(path, f"{base_name}{extension}")
    while os.path.exists(file_path):
        with open(file_path, "rb") as f: existing_content = f.read()
        if hashlib.sha256(existing_content).hexdigest() == file_hash: return True 
        counter += 1
        file_path = os.path.join(path, f"{base_name}_{counter}{extension}")
    with open(file_path, "wb") as f: f.write(payload)
    with open(HASH_REGISTRY, "a", newline="") as f:
        csv.writer(f).writerow([datetime.datetime.now(), bank_hint, rtype, use_date, file_hash])
    logging.info(f"SAVED: {file_path}")
    return True

def deep_scan_msg(msg_item, outlook, target_filter, received_tracker):
    try:
        current_sender = msg_item.SenderEmailAddress
        current_subject = msg_item.Subject
        base_bank_id = identify_bank(current_sender, current_subject)
        email_date = msg_item.ReceivedTime.date()
        for att in msg_item.Attachments:
            fname = att.FileName.lower()
            actual_bank = refine_bank_id(base_bank_id, fname)
            if fname.endswith((".xlsx", ".xls")):
                temp_path = os.path.join(QUARANTINE_DIR, "deep_" + fname)
                att.SaveAsFile(temp_path)
                with open(temp_path, "rb") as f: payload = f.read()
                rtype, rdate = analyze_excel(payload)
                if rtype != "UNKNOWN" and (target_filter == "ALL" or rtype == target_filter):
                    final_date = rdate if rdate else email_date
                    if save_payload(payload, actual_bank, rtype, final_date):
                        received_tracker[rtype].add(actual_bank)
                        print(f"    [+] Found {rtype} for {actual_bank}")
                try: os.remove(temp_path)
                except: pass
            elif fname.endswith(".msg"):
                temp_path = os.path.join(QUARANTINE_DIR, "nested_" + fname)
                att.SaveAsFile(temp_path)
                try:
                    nested_msg = outlook.Session.OpenSharedItem(temp_path)
                    deep_scan_msg(nested_msg, outlook, target_filter, received_tracker)
                    del nested_msg
                except: pass
                try: os.remove(temp_path)
                except: pass
    except: pass

# =========================================================
# 7. MAIN INTAKE (HYBRID MODE)
# =========================================================

def run_intake():
    print("\n" + "="*40)
    print(" RBZ RETURNS BOT - HYBRID MODE v5.0")
    print("="*40)

    print("\nWhich return type do you want to process?")
    print("1. Daily Loans (Single Date)")
    print("2. BSD 2/3 (Single Date)")
    print("3. BSD 4 (Date Range / Weekly)")
    print("4. ALL")
    choice = input("Select: ").strip()
    
    target_filter = "ALL"
    if choice == "1": target_filter = "DAILY_LOANS"
    elif choice == "2": target_filter = "BSD2_3"
    elif choice == "3": target_filter = "BSD4"

    if target_filter == "BSD4":
        print("\n[!] BSD 4 Selected: Please enter the week range.")
        while True:
            try:
                s_str = input("Start Date (YYYY-MM-DD): ").strip()
                e_str = input("End Date   (YYYY-MM-DD): ").strip()
                start_date = datetime.datetime.strptime(s_str, "%Y-%m-%d").date()
                end_date = datetime.datetime.strptime(e_str, "%Y-%m-%d").date()
                if start_date > end_date:
                    print("Start date cannot be after end date.")
                    continue
                break
            except: print("Invalid format. Use YYYY-MM-DD.")
    else:
        while True:
            user_date = input("\nScan emails received on (YYYY-MM-DD) or 'today': ").strip().lower()
            if user_date == 'today':
                d = datetime.date.today()
            else:
                try: d = datetime.datetime.strptime(user_date, "%Y-%m-%d").date()
                except: 
                    print("Invalid format.")
                    continue
            start_date = d
            end_date = d
            break

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)
    
    # --- METHOD A: FAST FILTER ---
    s_filter = start_date.strftime("%d/%m/%Y 00:00")
    e_filter = end_date.strftime("%d/%m/%Y 23:59")
    restriction = f"[ReceivedTime] >= '{s_filter}' AND [ReceivedTime] <= '{e_filter}'"
    
    print(f"\n[1] Trying Fast Filter ({s_filter} - {e_filter})...")
    
    target_messages = []
    try:
        filtered_messages = messages.Restrict(restriction)
        filtered_messages.Sort("[ReceivedTime]", True)
        
        # Verify if filter actually worked by checking dates manually on results
        for msg in filtered_messages:
            try:
                if start_date <= msg.ReceivedTime.date() <= end_date:
                    target_messages.append(msg)
            except: pass
    except: pass

    # --- METHOD B: FALLBACK (SAFETY NET) ---
    if len(target_messages) == 0:
        print("    ⚠️ Fast filter returned 0 results (Date Format Issue).")
        print("    [2] Switching to Manual Scan (100% Accuracy Mode)...")
        print("        Please wait, this scans newest emails until it hits your date...")
        
        target_messages = []
        count = 0
        scan_limit = 3000 # Safety limit
        
        for msg in messages:
            count += 1
            if count > scan_limit: 
                print("    Stopping scan (Limit Reached).")
                break
                
            try:
                msg_date = msg.ReceivedTime.date()
                # Optimized Exit: If we go PAST the start date (older), stop.
                if msg_date < start_date:
                    break
                
                if start_date <= msg_date <= end_date:
                    target_messages.append(msg)
            except: continue

    print(f"    Found {len(target_messages)} emails.")
    
    received_tracker = {k: set() for k in EXPECTED_RETURNS}
    for i, msg in enumerate(target_messages):
        deep_scan_msg(msg, outlook, target_filter, received_tracker)

    print("\n[+] Compliance Report:")
    rows = []
    cats = EXPECTED_RETURNS.keys() if target_filter == "ALL" else [target_filter]
    for rtype in cats:
        for b in EXPECTED_RETURNS[rtype]:
            status = "SUBMITTED" if b in received_tracker[rtype] else "MISSING"
            rows.append({"Type": rtype, "Bank": b, "Status": status})
            
    df = pd.DataFrame(rows)
    print(df)
    
    if not DRY_RUN:
        fname = f"Report_{start_date}_to_{end_date}.xlsx" if start_date != end_date else f"Report_{start_date}.xlsx"
        p = os.path.join(DASHBOARD_DIR, fname)
        df.to_excel(p, index=False)
        print(f"Saved: {p}")
    print("\nDone.")

if __name__ == "__main__":
    run_intake()