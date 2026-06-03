import win32com.client
import time
from datetime import datetime
import csv
import psutil

# List to store log messages
log_messages = []

def log_message(message, level="INFO"):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    full_message = f"[{timestamp}] [{level}] {message}"
    print(full_message)  # Terminal Print
    log_messages.append(full_message)  # Save for CSV

log_message(f"Process started.")

# Start a new Excel instance
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

# Open PERSONAL.XLSB from XLSTART
personal_path = r"C:\\Users\\SM775\\AppData\\Roaming\\Microsoft\\Excel\\XLSTART\\PERSONAL.XLSB"

try:
    workbook = excel.Workbooks.Open(personal_path)
except Exception as e:
    if "read-only" in str(e).lower():
        workbook = excel.Workbooks.Open(personal_path, ReadOnly=True)
    else:
        raise e

log_message(f"Microsoft Excel instance initiated in background mode")

# Run the macro
try:
    excel.Application.Run("PERSONAL.XLSB!CIC_AUTO")
    log_message("Macro executed successfully.")
except Exception as e:
    log_message(f"Macro failed: {e}")

# Optional: Close Excel
time.sleep(2)  # Wait to ensure macro finishes
workbook.Close(SaveChanges=True)
excel.Quit()

log_message(f"Excel instance used for macro closed")

# Kill background-only Excel instances
def close_background_excel_instances():
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] and 'EXCEL.EXE' in proc.info['name']:
            try:
                excel_app = win32com.client.GetObject(Class="Excel.Application")
                if not excel_app.Visible and excel_app.Workbooks.Count == 0:
                    proc.kill()
                    log_message(f"Killed background Excel process with PID {proc.pid}")
            except Exception:
                continue

close_background_excel_instances()

# Save log messages to a CSV file
csv_file_path = r"C:\\Users\\SM775\\Franke Group\\TM-HS-ESSC_IM.PL - Reports\\Python automation\\Logs\\CIC_log.csv"
with open(csv_file_path, mode='w', newline='') as file:
    log_message(f"Log messages saved to {csv_file_path}")
    writer = csv.writer(file)
    writer.writerow(["Log Messages"])
    for message in log_messages:
        writer.writerow([message])
