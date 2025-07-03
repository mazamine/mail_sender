# Mail Sender Application

Automatic mail sender from Excel file with GUI interface (French language support).

## Features

- 📧 Send personalized emails from Excel data
- 🎨 Modern GUI interface with preview
- 🔒 Secure email authentication using environment variables
- 📊 Excel file support (.xlsx, .xls)
- 🇫🇷 French language greetings (Monsieur/Madame)
- ⚡ Batch email processing

## Requirements

- Python 3.7+
- pandas
- openpyxl
- matplotlib
- tkinter (usually included with Python)

## Installation

1. Clone or download this repository
2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Setup Email Credentials

For security, set your email credentials as environment variables:

```bash
export SENDER_EMAIL="your_email@gmail.com"
export APP_PASSWORD="your_app_password"
```

### Getting Gmail App Password

1. Go to [Google Account Settings](https://myaccount.google.com/apppasswords)
2. Enable 2-factor authentication if not already enabled
3. Generate an app password for "Mail"
4. Use this password as your `APP_PASSWORD`

## Excel File Format

Your Excel file must contain these columns:
- `genre`: Gender ('m', 'f', 'monsieur', 'madame')
- `nom`: Last name
- `email`: Email address

Create a sample file:
```bash
python sample_contacts.py
```

## Usage

### GUI Mode (Recommended)
```bash
python main.py
```

### Command Line Mode
```bash
python gui.py
```

## Example Email Template

The application will automatically personalize emails:
- For `genre = 'm'` → "Bonjour Monsieur [nom]"
- For `genre = 'f'` → "Bonjour Madame [nom]"
- For other values → "Bonjour [nom]"

## Security Notes

- Never hardcode credentials in the source code
- Use environment variables for sensitive data
- Enable 2-factor authentication on your email account
- Use app-specific passwords for Gmail

## License

MIT License - see LICENSE file for details.
