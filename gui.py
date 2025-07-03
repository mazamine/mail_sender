import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from preview import render_latex_preview
from mailer import send_emails
import os

class MailSenderApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Mail Sender with LaTeX Preview")
        self.geometry("800x700")
        self.filename = None
        
        # Configure style
        self.configure(bg='#f0f0f0')

        self.create_widgets()

    def create_widgets(self):
        # Title
        title_label = tk.Label(self, text="Email Sender Application", 
                              font=("Arial", 16, "bold"), bg='#f0f0f0')
        title_label.pack(pady=10)
        
        # File selection frame
        file_frame = tk.Frame(self, bg='#f0f0f0')
        file_frame.pack(pady=10, padx=20, fill='x')
        
        tk.Label(file_frame, text="Excel File:", font=("Arial", 10, "bold"), 
                bg='#f0f0f0').pack(anchor='w')
        
        file_button_frame = tk.Frame(file_frame, bg='#f0f0f0')
        file_button_frame.pack(fill='x', pady=5)
        
        self.file_label = tk.Label(file_button_frame, text="No file selected", 
                                  bg='white', relief='sunken', anchor='w')
        self.file_label.pack(side='left', fill='x', expand=True, padx=(0, 10))
        
        tk.Button(file_button_frame, text="Browse", command=self.browse_file,
                 bg='#4CAF50', fg='white', font=("Arial", 10, "bold")).pack(side='right')

        # Email credentials info
        cred_frame = tk.Frame(self, bg='#f0f0f0')
        cred_frame.pack(pady=10, padx=20, fill='x')
        
        tk.Label(cred_frame, text="⚠️ Make sure to set SENDER_EMAIL and APP_PASSWORD environment variables", 
                font=("Arial", 9), fg='orange', bg='#f0f0f0').pack()

        # Subject field
        subject_frame = tk.Frame(self, bg='#f0f0f0')
        subject_frame.pack(pady=10, padx=20, fill='x')
        
        tk.Label(subject_frame, text="Email Subject:", font=("Arial", 10, "bold"), 
                bg='#f0f0f0').pack(anchor='w')
        self.subject_entry = tk.Entry(subject_frame, font=("Arial", 10))
        self.subject_entry.pack(fill='x', pady=5)

        # Body field
        body_frame = tk.Frame(self, bg='#f0f0f0')
        body_frame.pack(pady=10, padx=20, fill='both', expand=True)
        
        tk.Label(body_frame, text="Email Body:", font=("Arial", 10, "bold"), 
                bg='#f0f0f0').pack(anchor='w')
        self.body_text = scrolledtext.ScrolledText(body_frame, font=("Arial", 10), 
                                                  wrap=tk.WORD, height=12)
        self.body_text.pack(fill='both', expand=True, pady=5)

        # Buttons frame
        button_frame = tk.Frame(self, bg='#f0f0f0')
        button_frame.pack(pady=20)
        
        tk.Button(button_frame, text="Preview Email", command=self.preview_email,
                 bg='#2196F3', fg='white', font=("Arial", 10, "bold"), 
                 width=15).pack(side='left', padx=10)
        
        tk.Button(button_frame, text="Send Emails", command=self.send_all_emails,
                 bg='#FF5722', fg='white', font=("Arial", 10, "bold"), 
                 width=15).pack(side='left', padx=10)

        # Status label
        self.status_label = tk.Label(self, text="Ready", fg="blue", bg='#f0f0f0',
                                    font=("Arial", 10))
        self.status_label.pack(pady=10)

    def browse_file(self):
        self.filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if self.filename:
            filename_display = os.path.basename(self.filename)
            self.file_label.config(text=filename_display)
            self.status_label.config(text=f"Selected: {filename_display}", fg="green")

    def preview_email(self):
        subject = self.subject_entry.get()
        body = self.body_text.get("1.0", tk.END).strip()
        
        if not subject and not body:
            messagebox.showwarning("Warning", "Please enter a subject and body for the email.")
            return
        
        preview_text = f"Subject: {subject}\n\n{body}"
        render_latex_preview(preview_text)

    def send_all_emails(self):
        if not self.filename:
            messagebox.showerror("Error", "Please select an Excel file.")
            return

        subject = self.subject_entry.get().strip()
        body = self.body_text.get("1.0", tk.END).strip()
        
        if not subject:
            messagebox.showerror("Error", "Please enter an email subject.")
            return
            
        if not body:
            messagebox.showerror("Error", "Please enter an email body.")
            return

        # Check environment variables
        if not os.getenv("SENDER_EMAIL") or not os.getenv("APP_PASSWORD"):
            messagebox.showerror("Error", 
                               "Please set SENDER_EMAIL and APP_PASSWORD environment variables.\n\n"
                               "You can set them in your terminal:\n"
                               "export SENDER_EMAIL='your_email@gmail.com'\n"
                               "export APP_PASSWORD='your_app_password'")
            return

        try:
            self.status_label.config(text="Sending emails...", fg="orange")
            self.update()
            
            send_emails(self.filename, subject, body)
            self.status_label.config(text="✅ All emails sent successfully!", fg="green")
            messagebox.showinfo("Success", "All emails have been sent successfully!")
            
        except Exception as e:
            error_msg = str(e)
            self.status_label.config(text="❌ Error occurred", fg="red")
            messagebox.showerror("Error", f"Failed to send emails:\n\n{error_msg}")

if __name__ == "__main__":
    app = MailSenderApp()
    app.mainloop()
