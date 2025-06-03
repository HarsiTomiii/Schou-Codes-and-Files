# 📦 Maersk Stock Report – Excel Macro Runner

## 📋 Description

This Python script opens an Excel `.xlsm` workbook, runs a specified macro, saves the changes, and then closes the file.

It is tailored for internal use by Schou Company A/S for updating the MAERSK France stock report.

---

## ✨ How to Use

### 1. **Install Dependencies**

Ensure Python is installed with the `pywin32` library:

```bash
pip install pywin32
```

### 2. **Run the Script**

Set the correct macro-enabled Excel file and macro name in the script:

```python
file_path = r"C:\Users\tamhar\...\MAERSK_stock_report.xlsm"
macro_name = "refresh_all"
```

Then execute the script:

```bash
python main.py
```

### 3. **Macro Execution**

* The Excel file opens visibly.
* The `refresh_all` macro is executed.
* The workbook is saved and closed automatically.

---

## 🛠 Creating a Standalone Executable

Use `pyinstaller` to create a one-click executable:

```bash
pyinstaller --onefile --hidden-import win32com.client main.py
```

The resulting `.exe` file will be located in the `dist` directory.

---

## ⚠️ Notes

* Excel must be installed on the system running the script.
* Ensure Excel macro execution is allowed.
* Set `excel.Visible = False` in the script to hide the Excel UI during execution.
* File paths are hardcoded – update them to fit your environment.

---

## 🗂 Features

* Runs Excel macro automatically
* Saves workbook changes after execution
* Works with `.xlsm` macro-enabled files
* Can be converted to a standalone Windows executable

---

## 🔒 License

Internal use only – © Schou Company A/S
