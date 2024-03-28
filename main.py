
import tkinter as tk
import smtplib
import imaplib
import email
import email.header
from tkinter import ttk
import tkinter.messagebox
from tkhtmlview import HTMLLabel
import re
import datetime

class EmailClient(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Email Client")
        self.geometry("300x150")
        self.center_window()

        self.email_label = tk.Label(self, text="Email:")
        self.email_label.pack()

        self.email_entry = tk.Entry(self)
        self.email_entry.pack()

        self.password_label = tk.Label(self, text="Password:")
        self.password_label.pack()

        self.password_entry = tk.Entry(self, show="*")
        self.password_entry.pack()

        self.login_button = tk.Button(self, text="Login", command=self.login)
        self.login_button.pack()

    def center_window(self):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - 300) // 2
        y = (screen_height - 150) // 2
        self.geometry(f"300x150+{x}+{y}")
    def login(self):
        mail = self.email_entry.get()
        password = self.password_entry.get()

        smtp_server = smtplib.SMTP('smtp.gmail.com', 587)
        smtp_server.starttls()
        smtp_server.login(mail, password)

        imap_server = imaplib.IMAP4_SSL('imap.gmail.com', 993)
        imap_server.login(mail, password)
        imap_server.select('inbox')

        self.destroy()
        self.new_window = MainWindow(mail, password, smtp_server, imap_server)

class MainWindow(tk.Toplevel):
    def __init__(self, mail, password, smtp_server, imap_server, parent=None):
        super().__init__(parent)
        self.title("Main Window")
        self.geometry("1200x600")
        self.last_refresh_time = datetime.datetime.now()  # Inicjalizacja daty ostatniego odświeżenia

        self.mail = mail
        self.password = password
        self.smtp_server = smtp_server
        self.imap_server = imap_server
        self.loaded_messages = []
        self.read_messages = set()

        self.autoresponder_enabled = False
        self.autoresponder_message = "AUTOMATYCZNA ODPOWIEDZ: Jestem niedostepny, odpowiem pozniej."
        self.old_message_numbers = set()

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill=tk.BOTH)

        self.messages_container = ttk.Frame(self.notebook)
        self.notebook.add(self.messages_container, text="Messages")

        self.compose_button = tk.Button(self.messages_container, text="Compose Email", command=self.compose_email)
        self.compose_button.pack(side=tk.TOP, anchor='w', padx=10, pady=10)

        self.autoresponder_button = tk.Button(self.messages_container, text="Enable Autoresponder", command=self.toggle_autoresponder)
        self.autoresponder_button.pack(side=tk.TOP, anchor='w', padx=10, pady=10)

        self.messages_frame = ttk.Frame(self.messages_container)
        self.messages_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.keyword_label = tk.Label(self.messages_frame, text="Keyword:")
        self.keyword_label.pack(side=tk.TOP, anchor='w', padx=10)

        self.keyword_entry = tk.Entry(self.messages_frame)
        self.keyword_entry.pack(side=tk.TOP, anchor='w', padx=10)

        self.filter_button = tk.Button(self.messages_frame, text="Filter", command=self.filter_messages)
        self.filter_button.pack(side=tk.TOP, anchor='w', padx=10)

        self.messages_listbox = tk.Listbox(self.messages_frame, width=100, height=20)
        self.messages_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.email_content_html = HTMLLabel(self.messages_frame, html="")
        self.email_content_html.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.load_messages()
        self.refresh_emails()

    def refresh_emails(self):
        self.filter_messages()
        self.check_autoresponder()
        self.last_refresh_time = datetime.datetime.now()
        self.after(60000, self.refresh_emails)

    def filter_messages(self):
        keyword = self.keyword_entry.get().lower()

        self.messages_listbox.delete(0, tk.END)
        self.loaded_messages.clear()

        imap_server = self.imap_server
        imap_server.select('inbox')

        _, messages = imap_server.search(None, 'ALL')

        message_numbers = messages[0].split()
        start_index = max(0, len(message_numbers) - 40)  # Start index for slicing
        end_index = len(message_numbers)  # End index for slicing

        for num in reversed(message_numbers[start_index:end_index]):
            _, msg_data = imap_server.fetch(num, '(RFC822)')
            email_message = email.message_from_bytes(msg_data[0][1])
            sender = email_message['From']

            subject_encoded = email_message['Subject']
            subject_decoded = email.header.decode_header(subject_encoded)
            subject = ''
            for part, encoding in subject_decoded:
                if isinstance(part, bytes):
                    subject += part.decode(encoding or 'utf-8')
                else:
                    subject += part

            if keyword in subject.lower() or keyword in sender.lower():
                if num in self.read_messages:
                    message_info = f"From: {sender}, Subject: {subject}"
                else:
                    message_info = f"*From: {sender}, Subject: {subject}"

                self.messages_listbox.insert(tk.END, message_info)
                self.loaded_messages.append((num, email_message))

        self.messages_listbox.bind('<<ListboxSelect>>', self.display_email_content)

    def load_messages(self):
        imap_server = self.imap_server
        imap_server.select('inbox')

        _, messages = imap_server.search(None, 'ALL')

        message_numbers = messages[0].split()
        start_index = max(0, len(message_numbers) - 40)
        end_index = len(message_numbers)

        for num in reversed(message_numbers[start_index:end_index]):
            _, msg_data = imap_server.fetch(num, '(RFC822)')
            email_message = email.message_from_bytes(msg_data[0][1])
            sender = email_message['From']

            subject_encoded = email_message['Subject']
            subject_decoded = email.header.decode_header(subject_encoded)
            subject = ''
            for part, encoding in subject_decoded:
                if isinstance(part, bytes):
                    subject += part.decode(encoding or 'utf-8')
                else:
                    subject += part

            message_info = f"From: {sender}, Subject: {subject}"
            self.messages_listbox.insert(tk.END, message_info)
            self.loaded_messages.append((num, email_message))

    def display_email_content(self, event):
        selected_index = self.messages_listbox.curselection()
        if selected_index:
            selected_message_index = int(selected_index[0])
            selected_message_num, selected_message = self.loaded_messages[selected_message_index]

            self.read_messages.add(selected_message_num)

            payload = selected_message.get_payload()
            if isinstance(payload, str):
                content = payload
            else:
                content = ''
                for part in payload:
                    content += part.get_payload(decode=True).decode('utf-8', 'ignore')

            content = f'<div style="background-color: white;">{content}</div>'

            content = re.sub(r'color\s*:\s*(?!white|black)[^;]+;', '', content, flags=re.IGNORECASE)

            self.email_content_html.set_html(content)

    def compose_email(self):
        self.compose_window = tk.Toplevel(self)
        self.compose_window.title("Compose Email")
        self.compose_window.geometry("600x520")

        self.to_label = tk.Label(self.compose_window, text="To:")
        self.to_label.pack()

        self.to_entry = tk.Entry(self.compose_window)
        self.to_entry.pack()

        self.subject_label = tk.Label(self.compose_window, text="Subject:")
        self.subject_label.pack()

        self.subject_entry = tk.Entry(self.compose_window)
        self.subject_entry.pack()

        self.body_label = tk.Label(self.compose_window, text="Body:")
        self.body_label.pack()

        self.body_text = tk.Text(self.compose_window)
        self.body_text.pack()

        self.send_button = tk.Button(self.compose_window, text="Send", command=self.send_email)
        self.send_button.pack()

    def send_email(self):
        to = self.to_entry.get()
        subject = self.subject_entry.get()
        body = self.body_text.get("1.0", tk.END)

        message = f"Subject: {subject}\n\n{body}"

        smtp_server = self.smtp_server
        smtp_server.sendmail(self.mail, to, message)

        self.compose_window.destroy()
        tk.messagebox.showinfo("Email Sent", "Wiadomość wysłana pomyślnie.")

    def toggle_autoresponder(self):
        self.autoresponder_enabled = not self.autoresponder_enabled
        if self.autoresponder_enabled:
            self.autoresponder_button.config(text="Disable Autoresponder")
            self.old_message_numbers = set(num for num, _ in self.loaded_messages)
        else:
            self.autoresponder_button.config(text="Enable Autoresponder")

    def check_autoresponder(self):
        current_time = datetime.datetime.now()
        if self.autoresponder_enabled and (current_time - self.last_refresh_time).seconds >= 60:
            try:
                for num, msg in self.loaded_messages:
                    if num not in self.old_message_numbers:
                        sender_email = msg['From']
                        subject = msg['Subject']

                        smtp_server = smtplib.SMTP('smtp.gmail.com', 587)
                        smtp_server.starttls()
                        smtp_server.login(self.mail, self.password)
                        msg = f"From: {self.mail}\nSubject: Autoresponder: Re: {subject}\n\n{self.autoresponder_message}"
                        smtp_server.sendmail(self.mail, sender_email, msg)
                        smtp_server.quit()
            except Exception as e:
                print("Error sending autoresponder messages:", e)

if __name__ == "__main__":
    app = EmailClient()
    app.mainloop()