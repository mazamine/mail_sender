import pandas as pd
import smtplib
from email.mime.text import MIMEText

def send_emails(excel_path, subject, body_template):
    sender_email = "votremail@gmail.com"
    app_password = "votre_mot_de_passe_app"  # Généré depuis https://myaccount.google.com/apppasswords

    # Lire le fichier Excel
    df = pd.read_excel(excel_path)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender_email, app_password)

        for _, row in df.iterrows():
            genre = row['genre'].strip().lower()
            nom = row['nom'].strip()
            email = row['email'].strip()
            civilite = "Monsieur" if genre == "m" else "Madame"

            # Message personnalisé
            message_body = f"Bonjour {civilite} {nom},\n\n" + body_template

            msg = MIMEText(message_body, "plain", "utf-8")
            msg["Subject"] = subject
            msg["From"] = sender_email
            msg["To"] = email

            server.send_message(msg)
            print(f"✅ Email sent to {email}")
