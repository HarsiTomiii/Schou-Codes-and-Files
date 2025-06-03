# 📦 Maersk Forecast Update – Excel Macro Automation

## 📋 Description

This Python script automates the execution of an Excel macro inside a `.xlsm` workbook and exports the result as a timestamped `.xlsx` file in a predefined directory.

It is designed for internal use by Schou Company A/S to streamline the weekly container planning data export process.

---

## ✨ How to Use

### 1. **Install Dependencies**

Make sure you have Python installed with the required packages:

```bash
pip install pywin32
```

### 2. **Run the Script**

Update the script with your macro file path and macro name:

```python
file_path = r"C:\Users\tamhar\...\weekly_container_planning.xlsm"
macro_name = "refresh_all"
```

Then execute:

```bash
python main.py
```

### 3. **Macro Execution and Save**

* The script opens Excel and runs the specified macro.
* It then saves the result as an `.xlsx` file with a timestamp.
* The file is saved in:

```
C:\Users\tamhar\Schou Company A S\SC - INTERNT SUPPLY - Documents\General\Datafiles\container_planning_data
```

Example filename:

```
20250603123500_container_plan_raw.xlsx
```

---

## 🛠 Creating a Standalone Executable

Use `pyinstaller` to convert the script into a standalone `.exe` file:

```bash
pyinstaller --onefile --name maersk_forecast_update ^
--hidden-import win32com.client ^
--hidden-import win32com ^
--hidden-import win32com.client.dynamic ^
--hidden-import win32timezone ^
--hidden-import pythoncom ^
--hidden-import pywintypes ^
main.py
```

The executable will appear in the `/dist` directory.

---

## ⚠️ Notes

* **Excel must be installed** on the system running the script.
* Macro execution must be enabled in Excel.
* `excel.Visible = True` shows the Excel window — change to `False` to run hidden.
* The file paths are hardcoded and must be updated to match your environment.

---

## 🗂 Features

* Automates Excel macro execution
* Exports result as `.xlsx` with timestamp
* Silent execution with alerts disabled
* One-click .exe packaging

---

## 🔒 License

Internal use only – © Schou Company A/S
