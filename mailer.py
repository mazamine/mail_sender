import pandas as pd
import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def send_emails(excel_path, subject, body_template):
    """Send personalized emails to recipients from Excel file"""
    
    # Get email credentials from environment variables or use defaults
    sender_email = os.getenv("SENDER_EMAIL", "your_email@gmail.com")
    app_password = os.getenv("APP_PASSWORD", "your_app_password")
    
    if sender_email == "your_email@gmail.com" or app_password == "your_app_password":
        raise ValueError("Please set SENDER_EMAIL and APP_PASSWORD environment variables")

    try:
        # Read the Excel file
        df = pd.read_excel(excel_path)
        
        # Validate required columns
        required_columns = ['genre', 'nom', 'email']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise ValueError(f"Missing required columns in Excel file: {missing_columns}")

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender_email, app_password)

            for index, row in df.iterrows():
                try:
                    genre = str(row['genre']).strip().lower()
                    nom = str(row['nom']).strip()
                    email = str(row['email']).strip()
                    
                    # Skip rows with missing data
                    if not email or not nom or not genre:
                        print(f"⚠️ Skipping row {index + 1}: missing data")
                        continue
                    
                    # Determine civility
                    if genre in ['m', 'mr', 'monsieur']:
                        civilite = "Monsieur"
                    elif genre in ['f', 'mme', 'madame']:
                        civilite = "Madame"
                    else:
                        civilite = "Bonjour"  # Default greeting

                    # Create personalized message
                    message_body = f"Bonjour {civilite} {nom},\n\n{body_template}"

                    # Create email message
                    msg = MIMEMultipart()
                    msg["Subject"] = subject
                    msg["From"] = sender_email
                    msg["To"] = email
                    
                    msg.attach(MIMEText(message_body, "plain", "utf-8"))

                    server.send_message(msg)
                    print(f"✅ Email sent to {email}")
                    
                except Exception as e:
                    print(f"❌ Failed to send email to row {index + 1}: {str(e)}")
                    continue
                    
    except FileNotFoundError:
        raise ValueError(f"Excel file not found: {excel_path}")
    except Exception as e:
        raise ValueError(f"Error reading Excel file: {str(e)}")
