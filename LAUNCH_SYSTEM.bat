@echo off
TITLE RBZ System Launcher
echo Starting the RBZ Dashboard...
echo Please wait while the system initializes...

:: 1. Navigate to the current folder (where this file is)
cd /d "%~dp0"

:: 2. Force Streamlit to use a local config (Bypassing the Permission Error)
set STREAMLIT_CONFIG_DIR=%~dp0.streamlit

:: 3. Run the Dashboard
:: We use 'python' first, if that fails we try 'py'
python -m streamlit run dashboard.py --browser.gatherUsageStats false || py -m streamlit run dashboard.py --browser.gatherUsageStats false

pause