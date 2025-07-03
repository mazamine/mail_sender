import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from preview import render_latex_preview
from mailer import send_emails

class MailSenderApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Mail Sender with LaTeX Preview")
        self.geometry("800x600")
        self.filename = None

        self.create_widgets()

    def create_widgets(self):
        # Bouton pour sélectionner le fichier Excel
        tk.Button(self, text="Browse Excel File", command=self.browse_file).pack(pady=10)

        # Champ pour le sujet du mail
        tk.Label(self, text="Email Subject:").pack()
        self.subject_entry = tk.Entry(self, width=100)
        self.subject_entry.pack(pady=5)

        # Zone de texte pour le corps du mail
        tk.Label(self, text="Email Body (LaTeX allowed):").pack()
        self.body_text = scrolledtext.ScrolledText(self, width=100, height=15)
        self.body_text.pack(pady=5)

        # Boutons pour aperçu et envoi
        tk.Button(self, text="Preview Email", command=self.preview_email).pack(pady=10)
        tk.Button(self, text="Send Emails", command=self.send_all_emails).pack(pady=10)

        # Zone de statut
        self.status_label = tk.Label(self, text="", fg="blue")
        self.status_label.pack(pady=10)

    def browse_file(self):
        self.filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if self.filename:
            self.status_label.config(text=f"Selected file: {self.filename}")

    def preview_email(self):
        body = self.body_text.get("1.0", tk.END)
        render_latex_preview(body)

    def send_all_emails(self):
        if not self.filename:
            messagebox.showerror("Error", "Please select an Excel file.")
            return

        subject = self.subject_entry.get()
        body = self.body_text.get("1.0", tk.END)

        try:
            send_emails(self.filename, subject, body)
            self.status_label.config(text="✅ All emails sent successfully.")
        except Exception as e:
            messagebox.showerror("Error", str(e))
