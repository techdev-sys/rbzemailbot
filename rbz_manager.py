import os
import shutil
import datetime

# =========================================================
# CONFIGURATION
# =========================================================

ROOT_DIR = r"C:\RBZ_RETURNS_SYSTEM"
TEMPLATES_DIR = os.path.join(ROOT_DIR, "TEMPLATES")

# Map folder names to the specific Master File they need
TEMPLATE_MAP = {
    "DAILY_LOANS": "DAILY_LOANS_ANALYSIS.xlsx",
    "BSD2_3":      "MASTER_BSD23.xlsx",
    "BSD4":        "MASTER_BSD4.xlsx"
}

# =========================================================
# THE MANAGER LOGIC
# =========================================================

def deploy_masters():
    print("="*40)
    print(" RBZ FOLDER MANAGER - MASTER FILE DEPLOY")
    print("="*40)
    
    if not os.path.exists(TEMPLATES_DIR):
        print(f"[!] ERROR: Templates folder missing at {TEMPLATES_DIR}")
        return

    changes_made = 0

    # Walk through the entire RBZ system folder
    for root, dirs, files in os.walk(ROOT_DIR):
        
        # We only care about folders that look like Dates (e.g. "2026-01-26")
        folder_name = os.path.basename(root)
        try:
            # Check if folder name is a date YYYY-MM-DD
            datetime.datetime.strptime(folder_name, "%Y-%m-%d")
        except ValueError:
            continue # Not a date folder, skip it

        # Determine the Return Type by looking at the parent folders
        current_type = None
        if "DAILY_LOANS" in root: current_type = "DAILY_LOANS"
        elif "BSD2_3" in root:    current_type = "BSD2_3"
        elif "BSD4" in root:      current_type = "BSD4"
        
        if current_type and current_type in TEMPLATE_MAP:
            target_master = TEMPLATE_MAP[current_type]
            source_path = os.path.join(TEMPLATES_DIR, target_master)
            dest_path = os.path.join(root, target_master)
            
            # CHECK: Does the folder have the master file?
            if not os.path.exists(dest_path):
                if os.path.exists(source_path):
                    try:
                        shutil.copy2(source_path, dest_path)
                        print(f"[+] Deployed {target_master} to: {folder_name}")
                        changes_made += 1
                    except Exception as e:
                        print(f"[!] Error copying to {folder_name}: {e}")
                else:
                    print(f"[!] WARNING: Template {target_master} missing in TEMPLATES folder!")

    print("\n" + "="*40)
    if changes_made == 0:
        print(" All folders are up to date. No new Master Files needed.")
    else:
        print(f" DONE. Deployed {changes_made} Master Files.")
    print("="*40)

if __name__ == "__main__":
    deploy_masters()