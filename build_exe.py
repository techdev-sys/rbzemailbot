import PyInstaller.__main__
import os
import shutil

print("="*40)
print(" BUILDING RBZ PORTABLE SYSTEM")
print("="*40)

# 1. Create Wrapper
with open("run_app.py", "w") as f:
    f.write("""
import os
import sys
import streamlit.web.cli as stcli

if __name__ == "__main__":
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(__file__)
        
    os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
    os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
    os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"
    
    dashboard_path = os.path.join(base_path, 'dashboard.py')
    
    sys.argv = [
        "streamlit",
        "run",
        dashboard_path,
        "--global.developmentMode=false"
    ]
    sys.exit(stcli.main())
""")

# 2. Build Options
opts = [
    'run_app.py',
    '--name=RBZ_System',
    '--onedir',
    '--clean',
    '--noconfirm',
    
    '--add-data=dashboard.py;.',
    '--add-data=rbz_bot.py;.',
    
    # CRITICAL IMPORTS
    '--hidden-import=pandas',
    '--hidden-import=win32com',
    '--hidden-import=win32timezone',
    '--hidden-import=openpyxl',
    '--hidden-import=streamlit',
    
    '--collect-all=streamlit',
    '--collect-all=altair',
    '--collect-all=pandas',
]

PyInstaller.__main__.run(opts)

if os.path.exists("run_app.py"):
    os.remove("run_app.py")

print("\nBUILD COMPLETE!")