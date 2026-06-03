# 🧾 Python Excel Automation – Macro Runner

## 📌 Overview
This Python script automates the execution of an Excel macro (`CIC_AUTO`) stored in `PERSONAL.XLSB`. It runs Excel in the background, executes the macro, closes Excel, and saves execution logs to a CSV file.

---

## ⚙️ Features
- ✅ Runs Excel in invisible (background) mode  
- ✅ Executes a specific macro automatically  
- ✅ Handles read-only file access  
- ✅ Closes Excel safely after execution  
- ✅ Cleans up orphaned Excel processes  
- ✅ Saves detailed logs to CSV  

---

## 🧰 Requirements

### Python Libraries
Install dependencies:

```bash
pip install pywin32 psutil
