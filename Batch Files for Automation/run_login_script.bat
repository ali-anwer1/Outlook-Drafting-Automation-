@echo off
REM This batch file is used to run the login_script.py using the Python interpreter from a virtual environment.
REM Change directory to the location of the folder containing the Python Scripts.
cd /d "C:\path\Outlook-Drafting-Automation\Python Scripts"
REM Activate Python either by using the full path to your python.exe or by activating a virtual environment if you have one.
REM example: "C:\path\to\your\venv\Scripts\python.exe" or "C:\path\to\your\python.exe"
"path\python.exe" "login_script.py"
