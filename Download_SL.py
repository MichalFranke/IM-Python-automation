import subprocess
import pygetwindow as gw
import win32com.client
from datetime import datetime
import time
import csv
import os
import sys

# List to store log messages
log_messages = []

def log_message(message, level="INFO"):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    full_message = f"[{timestamp}] [{level}] {message}"
    print(full_message)
    log_messages.append(full_message)

log_message("Process started.")

sap_gui_path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\sapshcut.exe"
sap_system = 'PE1'
sap_client = '100'
sap_language = 'EN'

sap_command = [
    sap_gui_path,
    f'-system={sap_system}',
    f'-client={sap_client}',
    f'-language={sap_language}'
]

try:
    subprocess.run(["taskkill", "/F", "/IM", "saplogon.exe"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    subprocess.run(sap_command, check=True)
    time.sleep(5)
except subprocess.CalledProcessError as e:
    log_message(f"Failed to start SAP GUI: {e}", level="ERROR")
    sys.exit(1)

sap_windows = gw.getWindowsWithTitle('SAP')
if not sap_windows:
    log_message("SAP window not found.", level="ERROR")
    sys.exit(1)

sap_window = sap_windows[0]
sap_window.activate()
sap_window.maximize()
time.sleep(2)

try:
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application = sap_gui_auto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
except Exception as e:
    log_message(f"Failed to connect to SAP GUI scripting: {e}", level="ERROR")
    sys.exit(1)

log_message("SAP connection established.")

try:

    # Maximize main window
    session.findById("wnd[0]").maximize()

    # Navigate to transaction
    session.findById("wnd[0]/tbar[0]/okcd").text = "/n/FOX/PUR_SL_GEN"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]").sendVKey(17)

    # Fill popup fields
    session.findById("wnd[1]/usr/txtV-LOW").text = "SL DASH 2024"
    session.findById("wnd[1]/usr/txtENAME-LOW").text = "SM775"
    session.findById("wnd[1]").sendVKey(8)

    # Set today's date in YYYY-MM-DD format
    today = datetime.today().strftime("%Y-%m-%d")
    session.findById("wnd[0]/usr/ctxtS_EINDT-HIGH").text = today

    # Execute search
    session.findById("wnd[0]").sendVKey(8)

    # Export to Excel
    shell_path = "wnd[0]/usr/tabsTAB_CONTROL/tabpTAB_CONTROL_FC1/ssub0101_SCA:/FOX/PUR_SL_GEN:0101/cntlCUSTOM_0101/shellcont/shell"
    session.findById(shell_path).pressToolbarContextButton("&MB_EXPORT")
    session.findById(shell_path).selectContextMenuItem("&XXL")

    # Set export file name and path
    session.findById(
        "wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").text = "Current"
    session.findById(
        "wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").caretPosition = 7
    session.findById("wnd[1]").sendVKey(20)

    session.findById(
        "wnd[1]/usr/ctxtDY_PATH").text = r"C:\\Users\\SM775\\Franke Group\\TM-HS-ESSC_VENDORS.PL - ESSC-2-VENDORS - ESSC-2-VENDORS\\Service Level files\\"
    session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 103
    session.findById("wnd[1]").sendVKey(11)

    # Close transaction

except Exception as e:
    log_message(f"Error during SAP GUI automation: {e}", level="ERROR")
    sys.exit(1)

time.sleep(5)

csv_file_path = r"C:\Users\SM775\Franke Group\TM-HS-ESSC_IM.PL - Reports\Python automation\Logs\DWLD_SL_log.csv"
try:
    os.makedirs(os.path.dirname(csv_file_path), exist_ok=True)
    with open(csv_file_path, mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(["Log Messages"])
        for message in log_messages:
            writer.writerow([message])
    log_message(f"Log messages saved to {csv_file_path}")
except Exception as e:
    log_message(f"Failed to save log file: {e}", level="ERROR")
