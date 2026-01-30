import streamlit as st
import pandas as pd
import datetime
import os
import sys 
import json
import win32com.client
import pythoncom
import importlib 

# IMPORT YOUR BOT LOGIC
import rbz_bot 
from rbz_bot import deep_scan_msg, EXPECTED_RETURNS

# =========================================================
# CONFIGURATION
# =========================================================
st.set_page_config(page_title="RBZ Compliance Dashboard", layout="wide")

if getattr(sys, 'frozen', False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

USER_CONFIG_FILE = os.path.join(APP_DIR, "user_config.json")
DEFAULT_PATH = r"C:\RBZ_RETURNS_SYSTEM"

st.markdown("""
    <style>
    .big-font { font-size:20px !important; }
    .stButton>button { width: 100%; background-color: #004B87; color: white; }
    </style>
    """, unsafe_allow_html=True)

# =========================================================
# HELPER FUNCTIONS
# =========================================================

def get_outlook_accounts():
    accounts = []
    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        for folder in namespace.Folders:
            accounts.append(folder.Name)
    except: pass
    return accounts

def load_user_config():
    if os.path.exists(USER_CONFIG_FILE):
        try:
            with open(USER_CONFIG_FILE, "r") as f:
                return json.load(f)
        except: pass
    return None

def save_user_config(account_name, save_path):
    save_path = os.path.normpath(save_path)
    config_data = {
        "account_name": account_name, 
        "save_path": save_path,
        "registered_at": str(datetime.datetime.now())
    }
    with open(USER_CONFIG_FILE, "w") as f:
        json.dump(config_data, f)

def get_status_icon(status):
    return "✅" if status == "SUBMITTED" else "❌"

# =========================================================
# SCAN LOGIC (GOLD MASTER: INFINITE INDEX WALK)
# =========================================================

def run_scan_logic(return_type, start_date, end_date):
    pythoncom.CoInitialize()
    
    # 1. SETUP
    user_config = load_user_config()
    target_account = user_config.get('account_name', '') if user_config else ''
    save_path = user_config.get('save_path', DEFAULT_PATH) if user_config else DEFAULT_PATH
    
    rbz_bot.update_save_path(save_path)
    
    status_text = st.empty()
    progress_bar = st.progress(0)
    
    status_text.text(f"Connecting to {target_account}...")
    
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # 2. CONNECT TO INBOX
        try:
            inbox = namespace.Folders[target_account].Folders["Inbox"]
        except:
            st.error(f"❌ Could not access Inbox for '{target_account}'.")
            return None

        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True) # Sort Descending (Newest = Item 1)
        
        # 3. INFINITE INDEX SCAN (EXE SAFE)
        target_messages = []
        
        # Safety limit of 30,000 prevents freezing, but logic stops well before this.
        total_items = messages.Count
        safety_limit = min(total_items, 30000) 
        
        status_text.text(f"Searching email history ({start_date} to {end_date})...")
        
        found_count = 0
        
        # Outlook uses 1-based indexing
        for i in range(1, safety_limit + 1):
            
            # Show activity every 100 items
            if i % 100 == 0:
                status_text.text(f"Checked {i} emails... (Found {found_count} matches so far)")

            try:
                # DIRECT GRAB (EXE Safe)
                msg = messages.Item(i)
                msg_date = msg.ReceivedTime.date()
                
                # A. TOO NEW? Skip and keep walking down.
                if msg_date > end_date:
                    continue 
                
                # B. TOO OLD? Stop immediately.
                if msg_date < start_date:
                    status_text.text(f"Reached emails older than {start_date}. Stopping.")
                    break
                
                # C. MATCH? Keep it.
                if start_date <= msg_date <= end_date:
                    target_messages.append(msg)
                    found_count += 1
                    
            except: continue

        # 4. PROCESS RESULTS
        total = len(target_messages)
        
        if total == 0:
            st.warning(f"⚠️ No emails found between {start_date} and {end_date}.")
            return None
            
        status_text.text(f"Found {total} emails. Extracting files...")
        
        received_tracker = {k: set() for k in EXPECTED_RETURNS}
        
        for i, msg in enumerate(target_messages):
            progress_bar.progress(int(((i + 1) / total) * 100))
            try:
                rbz_bot.deep_scan_msg(msg, outlook, return_type, received_tracker)
            except: pass
            
        progress_bar.progress(100)
        status_text.success(f"Scan Complete! Processed {total} emails.")
        return received_tracker

    except Exception as e:
        st.error(f"❌ System Error: {e}")
        return None

# =========================================================
# MAIN APP
# =========================================================

def main():
    user_config = load_user_config()

    # SELF HEALING
    if user_config and 'account_name' not in user_config:
        user_config = None
    
    # --- PHASE 1: MAGIC SETUP WIZARD ---
    if user_config is None:
        st.title("🏦 RBZ System Setup")
        st.info("Welcome! Let's connect the bot to your Outlook.")
        
        available_accounts = get_outlook_accounts()
        
        if not available_accounts:
            st.error("❌ Could not detect any Outlook accounts. Please open Outlook and try again.")
            if st.button("Retry"): st.rerun()
            return

        col1, col2 = st.columns(2)
        with col1:
            selected_account = st.selectbox("Select which Outlook Inbox to scan:", available_accounts)
        with col2:
            path_in = st.text_input("Where should files be saved?", value=DEFAULT_PATH)
        
        st.write("---")
        if st.button("💾 Connect & Launch"):
            try:
                os.makedirs(path_in, exist_ok=True)
                save_user_config(selected_account, path_in)
                st.success("Connected! Launching Dashboard...")
                st.rerun()
            except Exception as e:
                st.error(f"Error creating folder: {e}")
        return

    # --- PHASE 2: DASHBOARD ---
    current_account = user_config.get('account_name', 'Unknown')
    current_path = user_config.get('save_path', DEFAULT_PATH)
    
    st.title("🏦 RBZ Compliance Dashboard")
    st.markdown("---")

    with st.sidebar:
        st.header("⚙️ Settings")
        st.caption(f"📧 Scanning: **{current_account}**")
        st.caption(f"📂 Saving to: {current_path}")
        
        if st.button("Change Account / Reset"):
            try: os.remove(USER_CONFIG_FILE)
            except: pass
            st.rerun()
        
        st.divider()
        
        mode = st.radio("Return Type:", ["Daily Loans", "BSD 2/3", "BSD 4 (Weekly)"])
        type_map = {"Daily Loans": "DAILY_LOANS", "BSD 2/3": "BSD2_3", "BSD 4 (Weekly)": "BSD4"}
        selected_type = type_map[mode]
        
        if mode == "BSD 4 (Weekly)":
            today = datetime.date.today()
            start_d = st.date_input("Start (Mon)", today - datetime.timedelta(days=today.weekday()))
            end_d = st.date_input("End (Fri)", start_d + datetime.timedelta(days=4))
        else:
            scan_date = st.date_input("Scan Date", datetime.date.today())
            start_d = scan_date
            end_d = scan_date
            
        st.markdown("---")
        if st.button("🚀 START EXTRACTION", type="primary"):
            with st.spinner("Extracting..."):
                tracker = run_scan_logic(selected_type, start_d, end_d)
                if tracker:
                    st.session_state['tracker'] = tracker
                    st.session_state['scan_done'] = True
                    st.session_state['run_type'] = selected_type

    if 'scan_done' in st.session_state and st.session_state['scan_done']:
        tracker = st.session_state['tracker']
        rtype = st.session_state['run_type']
        
        bank_list = EXPECTED_RETURNS[rtype]
        rows = []
        submitted = 0
        
        for bank in bank_list:
            status = "SUBMITTED" if bank in tracker[rtype] else "MISSING"
            if status == "SUBMITTED": submitted += 1
            
            rows.append({
                "Bank Name": bank,
                "Status": status,
                "Check": get_status_icon(status)
            })
            
        df = pd.DataFrame(rows)
        
        c1, c2, c3 = st.columns(3)
        c1.metric("Total", len(bank_list))
        c2.metric("Received", submitted)
        c3.metric("Pending", len(bank_list) - submitted, delta_color="inverse")
        
        st.dataframe(
            df[["Bank Name", "Check", "Status"]],
            width="stretch",
            hide_index=True,
            column_config={
                "Status": st.column_config.TextColumn(
                    "Compliance Status",
                    validate="^(SUBMITTED|MISSING)$"
                ),
                "Check": st.column_config.TextColumn("Verdict")
            }
        )
        st.success(f"Files saved to: {os.path.join(current_path, rtype)}")

if __name__ == "__main__":
    main()