# Bulk Email Sender with Attachment Support (GUI-Based)

## 📋 Description

This Python application provides a graphical user interface (GUI) for sending bulk emails using a template and recipient list from an Excel file. It allows the user to:

* Log in with an email account (supports Office365 SMTP).
* Select an HTML email template.
* Select an Excel file with recipient email addresses.
* Optionally attach a file.
* Review the selected files and email subject before sending.
* Send personalized emails to all recipients listed in the Excel file.

Each email is sent with the same subject and body, and optionally includes an attachment.

---

## ✨ How to Use

### **Run the Program**

Run the Python script:

```bash
python bulk_email_sender.py
```
Or you can run the executable:

```bash
output/mass_email.exe
```
### 3. **Login**

* Enter your **email address**, **password**, and the **From** email (typically same as login email).
* Click **Login** to proceed.

### 4. **Select Files**

* **HTML Template File**: The body of the email will use this HTML content.
* **Excel File**: Should contain recipient emails in the **first column** starting from row 1.
* **Attachment File** (optional): Any file you wish to send as an attachment with each email.

### 5. **Set Subject**

You'll be prompted to enter the email subject line.

### 6. **Confirm and Send**

A confirmation window summarizes your selections.

* Click **Send emails** to begin sending.
* A popup will notify you once all emails have been successfully sent.

---

## ✅ Input File Format

### 📄 Excel File (`.xlsx`)

The Excel file must have the email addresses listed in **column A**, one per row. Example:

| Email Address                                 |
| --------------------------------------------- |
| [user1@example.com](mailto:user1@example.com) |
| [user2@example.com](mailto:user2@example.com) |
| ...                                           |

### 📄 HTML Template

Any valid `.html` or `.htm` file. This will be used as the body of the email.
You can create and save an email into a `.html` or `.htm` file and use that for the template.

---

## 💡 Outcome

Once completed:

* All recipients from the Excel sheet will receive an email.
* Each email will have the same subject, HTML body, and optional attachment.
* A message will display the number of successfully sent emails or show an error if something goes wrong.

---

## 📧 Notes

* Uses `smtp.office365.com:587` for sending. You must use a Microsoft 365 email account.
* Consider using an [App Password](https://support.microsoft.com/account-billing/create-app-passwords-to-use-apps-that-don-t-use-two-step-verification-73033404-8a9c-4e8c-9013-1c8e9d6c9926) if Multi-Factor Authentication (MFA) is enabled on your account.
* Ensure less secure app access is enabled if required by your provider.

---

## 🛦 Features

* Simple and intuitive GUI with Tkinter
* Multi-file selection dialog
* Email template preview
* Attachment support
* SMTP authentication
* Excel-based bulk email support
* Success and error messages
* create executable with `auto-py-to-exe` command

---

## 📌 License

This project is provided as-is without any warranties. You are free to use, modify, and distribute it.
