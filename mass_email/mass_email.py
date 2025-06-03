import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib
import os
import openpyxl


def send_emails(email, password, from_email, subject, template_file, excel_file, attachment_file):
    try:
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()
        server.login(email, password)

        wb = openpyxl.load_workbook(excel_file)
        sheet = wb.active

        with open(template_file, 'r', encoding='utf-8', errors='ignore') as f:
            template_content = f.read()

        recipients_count = 0

        for row in sheet.iter_rows(min_row=1, max_col=1, values_only=True):
            if row[0]:
                recipient = row[0]
                msg = MIMEMultipart()
                msg['From'] = from_email
                msg['To'] = recipient
                msg['Subject'] = subject

                body = template_content
                msg.attach(MIMEText(body, 'html'))

                # Attachments
                if attachment_file:
                    attachment = open(attachment_file, "rb")
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(attachment_file)}")
                    msg.attach(part)
                    attachment.close()

                server.sendmail(from_email, recipient, msg.as_string())
                recipients_count += 1

        server.quit()
        messagebox.showinfo("Success", f"Emails sent successfully to {recipients_count} recipients")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")


def confirm_send(email, password, from_email, subject, template_file, excel_file, attachment_file):
    confirm_window = tk.Tk()
    confirm_window.title("Confirmation")
    confirm_window.geometry("400x220")

    tk.Label(confirm_window, text="Selected Template File: ").grid(row=0, column=0, sticky="w")
    tk.Label(confirm_window, text=os.path.basename(template_file)).grid(row=0, column=1, sticky="w")

    tk.Label(confirm_window, text="Selected Excel File: ").grid(row=1, column=0, sticky="w")
    tk.Label(confirm_window, text=os.path.basename(excel_file)).grid(row=1, column=1, sticky="w")

    tk.Label(confirm_window, text="Attachment File: ").grid(row=2, column=0, sticky="w")
    tk.Label(confirm_window, text=os.path.basename(attachment_file) if attachment_file else "None").grid(row=2,
                                                                                                         column=1,
                                                                                                         sticky="w")

    tk.Label(confirm_window, text="Email Subject: ").grid(row=3, column=0, sticky="w")
    tk.Label(confirm_window, text=subject).grid(row=3, column=1, sticky="w")

    tk.Label(confirm_window, text="From Email: ").grid(row=4, column=0, sticky="w")
    tk.Label(confirm_window, text=from_email).grid(row=4, column=1, sticky="w")

    def send():
        confirm_window.destroy()
        send_emails(email, password, from_email, subject, template_file, excel_file, attachment_file)

    def cancel():
        confirm_window.destroy()

    tk.Button(confirm_window, text="Send emails", command=send).grid(row=5, column=0, padx=10, pady=20)
    tk.Button(confirm_window, text="Cancel", command=cancel).grid(row=5, column=1, padx=10, pady=20)

    confirm_window.mainloop()


def select_files(email, password, from_email):
    template_file = filedialog.askopenfilename(title="Select template file",
                                               filetypes=(("HTML files", "*.htm;*.html"), ("All files", "*.*")))
    if not template_file:
        messagebox.showerror("Error", "No template file selected")
        return

    excel_file = filedialog.askopenfilename(title="Select Excel file",
                                            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
    if not excel_file:
        messagebox.showerror("Error", "No Excel file selected")
        return

    attachment_file = filedialog.askopenfilename(title="Select attachment file",
                                                 filetypes=(("All files", "*.*"),))

    subject = simpledialog.askstring("Subject", "Enter email subject:")
    if subject is not None:
        confirm_send(email, password, from_email, subject, template_file, excel_file, attachment_file)


def login_window():
    login_window = tk.Tk()
    login_window.title("Login")
    login_window.geometry("300x220")

    tk.Label(login_window, text="Email:").grid(row=0, column=0, sticky="w")
    tk.Label(login_window, text="Password:").grid(row=1, column=0, sticky="w")
    tk.Label(login_window, text="From Email:").grid(row=2, column=0, sticky="w")

    email_entry = tk.Entry(login_window)
    password_entry = tk.Entry(login_window, show="*")
    from_email_entry = tk.Entry(login_window)

    email_entry.grid(row=0, column=1)
    password_entry.grid(row=1, column=1)
    from_email_entry.grid(row=2, column=1)

    def send_credentials():
        email = email_entry.get()
        password = password_entry.get()
        from_email = from_email_entry.get()
        login_window.destroy()
        select_files(email, password, from_email)

    tk.Button(login_window, text="Login", command=send_credentials).grid(row=3, column=0, columnspan=2, pady=10)

    login_window.mainloop()


def main():
    login_window()


if __name__ == "__main__":
    main()