
import os
import smtplib
import traceback
import requests
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv
import re
import time
import hashlib
import pickle
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

from dotenv import dotenv_values

# Load environment variables from .env file
env_vars = dotenv_values(".env")
EMAIL_SENDER = env_vars.get("EMAIL_SENDER")
EMAIL_PASSWORD = env_vars.get("EMAIL_PASSWORD")
SMTP_SERVER = env_vars.get("SMTP_SERVER")
SMTP_PORT = int(env_vars.get("SMTP_PORT", 587))
ADMIN_EMAILS_STR = env_vars.get("ADMIN_EMAILS")
if ADMIN_EMAILS_STR:
    ADMIN_EMAILS = [email.strip() for email in ADMIN_EMAILS_STR.split(",")]
else:
    ADMIN_EMAILS = []

# --- Telegram Configuration ---
TELEGRAM_BOT_TOKEN = env_vars.get("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = env_vars.get("TELEGRAM_CHAT_ID")

# --- Email validation setup ---
EMAIL_REGEX = re.compile(
    r"^[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@"
    r"[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?(?:\.[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?)*$",
    re.IGNORECASE,
)

# Try to import dnspython for MX lookups; make MX check optional if not installed
try:
    import dns.resolver
    MX_CHECK_AVAILABLE = True
except Exception:
    MX_CHECK_AVAILABLE = False

def clean_email_raw(value):
    """Apply PROPER(), TRIM(), LOWER() pipeline and remove control chars."""
    if value is None:
        return None
    s = str(value)
    # PROPER (title case) as requested, then TRIM, then LOWER
    s = s.title()
    s = s.strip()
    s = s.lower()
    # remove non-printable/control characters
    s = ''.join(ch for ch in s if ch.isprintable())
    return s

def is_valid_email_format(email):
    if not email or '@' not in email:
        return False
    # quick checks
    if email.count('@') != 1:
        return False
    local, domain = email.rsplit('@', 1)
    if not local or not domain:
        return False
    # domain must have at least one dot
    if '.' not in domain:
        return False
    # regex check
    return bool(EMAIL_REGEX.match(email))

def has_mx_record(domain):
    if not MX_CHECK_AVAILABLE:
        return True
    try:
        answers = dns.resolver.resolve(domain, 'MX')
        return len(answers) > 0
    except Exception:
        return False

def get_file_hash(filepath):
    """Compute SHA256 hash of the file."""
    hash_sha256 = hashlib.sha256()
    with open(filepath, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_sha256.update(chunk)
    return hash_sha256.hexdigest()

def load_cache():
    """Load cached data if exists."""
    if os.path.exists(CACHE_FILE_PATH):
        with open(CACHE_FILE_PATH, 'rb') as f:
            return pickle.load(f)
    return None

def save_cache(hash_val, df, analytics):
    """Save hash and cleaned data to cache."""
    with open(CACHE_FILE_PATH, 'wb') as f:
        pickle.dump({'hash': hash_val, 'df': df, 'analytics': analytics}, f)

SENT_LOG_FILE_PATH = "sent_log.pkl"

def load_sent_log():
    """Load sent log if exists."""
    if os.path.exists(SENT_LOG_FILE_PATH):
        with open(SENT_LOG_FILE_PATH, 'rb') as f:
            return pickle.load(f)
    return {}

def save_sent_log(log):
    """Save sent log."""
    with open(SENT_LOG_FILE_PATH, 'wb') as f:
        pickle.dump(log, f)

def validate_email_entry(val, invalid_emails):
    if not val or str(val).lower() in ['nan', 'none', '']:
        invalid_emails.append(str(val))
        return []
    # Split by comma
    emails = [e.strip() for e in str(val).split(',') if e.strip()]
    valid_emails = []
    for email in emails:
        cleaned = clean_email_raw(email)
        if not cleaned:
            continue
        original_cleaned = cleaned
        if not is_valid_email_format(cleaned):
            # Try correction
            corrected = attempt_email_correction(cleaned)
            if corrected and is_valid_email_format(corrected):
                domain = corrected.rsplit('@', 1)[-1]
                if has_mx_record(domain):
                    print(f"Corrected invalid email '{email}' to '{corrected}'")
                    cleaned = corrected
                else:
                    continue
            else:
                continue
        # MX check (optional)
        domain = cleaned.rsplit('@', 1)[-1]
        if not has_mx_record(domain):
            continue
        valid_emails.append(cleaned)
    if not valid_emails:
        invalid_emails.append(str(val))
    return valid_emails

def attempt_email_correction(email):
    """Attempt to correct common email errors."""
    if not email:
        return None
    # Remove extra @ symbols, keep the last part
    parts = email.split('@')
    if len(parts) > 2:
        # Assume the first is local, last is domain
        local = '@'.join(parts[:-1])
        domain = parts[-1]
        corrected = f"{local}@{domain}"
    else:
        corrected = email
    # Fix common typos
    corrected = corrected.replace(',com', '.com').replace(' ', '').replace('.com.', '.com').replace(', ', ',')
    # Remove trailing commas or dots
    corrected = corrected.rstrip(',. ')
    # If domain missing dot, add .com
    if '@' in corrected and '.' not in corrected.split('@')[1]:
        corrected = corrected + '.com'
    # If no change, return None
    return corrected if corrected != email else None

def retry_sendmail(server, sender, recipients, msg, max_retries=3, backoff=2):
    """Retry SMTP sendmail with exponential backoff."""
    for attempt in range(max_retries):
        try:
            server.sendmail(sender, recipients, msg)
            return True
        except Exception as e:
            if attempt < max_retries - 1:
                print(f"Sendmail attempt {attempt + 1} failed: {e}. Retrying in {backoff} seconds...")
                time.sleep(backoff)
                backoff *= 2  # Exponential backoff
            else:
                raise e
    return False

# --- File Paths ---
# Make sure these files exist in the same directory as the script
EXCEL_FILE_PATH = "December.xlsx"
BIRTHDAY_IMAGE_PATH = "birthday.png"
CACHE_FILE_PATH = "email_cache.pkl"

# --- Email Content ---
EMAIL_SUBJECT = "Cheers to You on Your Special Day! 🥂🎈 - NCS Wishes 🎂"
EMAIL_BODY = "Wishing you a very happy birthday!"

def send_birthday_email(recipient_to, recipients_bcc, subject, body, image_path):
    """
    Sends the birthday email with an image attachment.
    """
    if not all([EMAIL_SENDER, EMAIL_PASSWORD, SMTP_SERVER]):
        raise Exception("Error: Email sender configuration is incomplete in .env file.")

    msg = MIMEMultipart()
    msg["From"] = EMAIL_SENDER
    msg["To"] = recipient_to
    msg["Subject"] = subject

    if recipients_bcc:
        msg["Bcc"] = ", ".join(recipients_bcc)

    msg.attach(MIMEText(body, "plain"))

    try:
        with open(image_path, "rb") as f:
            img = MIMEImage(f.read())
            img.add_header("Content-Disposition", "attachment", filename=os.path.basename(image_path))
            msg.attach(img)
    except FileNotFoundError:
        raise Exception(f"Error: Birthday image not found at {image_path}")

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(EMAIL_SENDER, EMAIL_PASSWORD)
        all_recipients = [recipient_to] + recipients_bcc
        retry_sendmail(server, EMAIL_SENDER, all_recipients, msg.as_string())

def send_notification_email(subject, body):
    """
    Sends a plain text notification email to the admins.
    """
    if not ADMIN_EMAILS:
        print("Warning: ADMIN_EMAILS not set. Skipping email notification.")
        return

    if not all([EMAIL_SENDER, EMAIL_PASSWORD, SMTP_SERVER]):
        print("Error: Cannot send email notification. Email sender configuration is incomplete.")
        return
        
    msg = MIMEMultipart()
    msg["From"] = EMAIL_SENDER
    msg["To"] = ", ".join(ADMIN_EMAILS)
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            retry_sendmail(server, EMAIL_SENDER, ADMIN_EMAILS, msg.as_string())
            print(f"Successfully sent email notification to {', '.join(ADMIN_EMAILS)}")
    except Exception as e:
        print(f"Failed to send email notification: {e}")

def send_telegram_notification(message):
    """
    Sends a message to the admin via a Telegram bot.
    """
    if not all([TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID]):
        print("Warning: Telegram credentials not set. Skipping Telegram notification.")
        return
    
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": message,
        "parse_mode": "Markdown"
    }
    
    try:
        response = requests.post(url, json=payload)
        if response.status_code == 200:
            print(f"Successfully sent Telegram notification.")
        else:
            print(f"Failed to send Telegram notification: {response.text}")
    except Exception as e:
        print(f"Failed to send Telegram notification: {e}")


def check_birthdays_and_send_email():
    """
    Checks for birthdays, sends emails, and returns a status message.
    Raises an exception on critical errors.
    """
    analytics = {
        'total_rows': 0,
        'valid_emails': 0,
        'birthdays_found': 0,
        'emails_sent': 0,
        'send_failures': 0,
        'corrections': [],
        'invalid_emails': [],
        'skipped_sends': 0
    }
    
    # Check cache
    current_hash = get_file_hash(EXCEL_FILE_PATH)
    cache = load_cache()
    if cache and cache['hash'] == current_hash:
        df = cache['df']
        analytics = cache['analytics']
        print("Using cached cleaned data (file unchanged).")
        # Ensure new keys are present
        analytics.setdefault('corrections', [])
        analytics.setdefault('invalid_emails', [])
        analytics.setdefault('skipped_sends', 0)
    else:
        print("File changed or no cache; validating emails...")
        try:
            df = pd.read_excel(EXCEL_FILE_PATH)
            analytics['total_rows'] = len(df)
        except FileNotFoundError:
            raise Exception(f"Error: Excel file not found at {EXCEL_FILE_PATH}")
        except Exception as e:
            raise Exception(f"Error reading Excel file: {e}")

        # --- Data Validation and Cleaning ---
        # Clean column headers: remove whitespace and convert to uppercase
        df.columns = df.columns.str.strip().str.upper()

        # Now, check for the cleaned column names
        if "EMAIL" not in df.columns or "DOB" not in df.columns:
            raise Exception("Error: Excel file must contain 'EMAIL' and 'DOB' columns.")

        # Validate and clean 'EMAIL' column using strict rules
        initial_count = len(df)
        invalid_emails = []

        def validate_email_entry(val, invalid_emails):
            if not val or str(val).lower() in ['nan', 'none', '']:
                invalid_emails.append(str(val))
                return []
            # Split by comma
            emails = [e.strip() for e in str(val).split(',') if e.strip()]
            valid_emails = []
            for email in emails:
                cleaned = clean_email_raw(email)
                if not cleaned:
                    continue
                original_cleaned = cleaned
                if not is_valid_email_format(cleaned):
                    # Try correction
                    corrected = attempt_email_correction(cleaned)
                    if corrected and is_valid_email_format(corrected):
                        # domain = corrected.rsplit('@', 1)[-1]
                        # if has_mx_record(domain):
                        analytics['corrections'].append(f"Corrected '{email}' to '{corrected}'")
                        cleaned = corrected
                        # else:
                        #     continue
                    else:
                        continue
                # MX check (optional)
                # domain = cleaned.rsplit('@', 1)[-1]
                # if not has_mx_record(domain):
                #     continue
                valid_emails.append(cleaned)
            if not valid_emails:
                invalid_emails.append(str(val))
            return valid_emails

        df['EMAIL'] = df['EMAIL'].apply(lambda x: validate_email_entry(x, invalid_emails))
        # Drop rows without valid emails
        df = df[df['EMAIL'].apply(len) > 0]
        analytics['valid_emails'] = len(df)
        analytics['invalid_emails'] = invalid_emails
        
        # Save to cache
        save_cache(current_hash, df, analytics)
    
    # --- Birthday Checking ---
    df['DOB'] = pd.to_datetime(df['DOB'], errors='coerce')
    df = df.dropna(subset=['DOB'])
    
    today = pd.Timestamp.today()
    birthdays_today = df[df['DOB'].dt.month.eq(today.month) & df['DOB'].dt.day.eq(today.day)]
    analytics['birthdays_found'] = len(birthdays_today)
    
    if birthdays_today.empty:
        print("No birthdays today.")
        return "No birthdays today.", analytics
    
    # Load sent log
    date_str = today.strftime('%Y-%m-%d')
    sent_log = load_sent_log()
    sent_today = sent_log.get(date_str, set())
    
    # Send emails
    for index, row in birthdays_today.iterrows():
        emails = row['EMAIL']
        for email in emails:
            if email not in sent_today:
                try:
                    send_birthday_email(email, [], EMAIL_SUBJECT, EMAIL_BODY, BIRTHDAY_IMAGE_PATH)
                    analytics['emails_sent'] += 1
                    sent_today.add(email)
                except Exception as e:
                    print(f"Failed to send email to {email}: {e}")
                    analytics['send_failures'] += 1
            else:
                analytics['skipped_sends'] += 1
    
    # Save sent log
    sent_log[date_str] = sent_today
    save_sent_log(sent_log)
    
    return f"Successfully processed {analytics['emails_sent']} new birthday emails.", analytics

if __name__ == "__main__":
    try:
        status_message, analytics = check_birthdays_and_send_email()
        print(status_message)
        
        # Create detailed report
        report = f"{status_message}\n\n--- Analytics ---\n"
        for k, v in analytics.items():
            if k in ['corrections', 'invalid_emails']:
                report += f"{k.replace('_', ' ').title()}: {len(v)}\n"
            else:
                report += f"{k.replace('_', ' ').title()}: {v}\n"
        
        if analytics['corrections']:
            report += "\n--- Corrections ---\n" + "\n".join(analytics['corrections']) + "\n"
        if analytics['invalid_emails']:
            report += "\n--- Invalid Emails ---\n" + "\n".join(f"- {e}" for e in analytics['invalid_emails']) + "\n"
        
        # Send notifications
        send_notification_email("Birthday Cron Job: Daily Report", report)
        send_telegram_notification(f"*Birthday Cron Job: Daily Report*\n\n{report}")

    except Exception as e:
        error_message = f"Birthday Cron Job: FAILED\n\nError: {e}"
        # For Telegram, send a shorter message without the full traceback for readability
        telegram_error_message = f"*Birthday Cron Job: FAILED*\n\n`{e}`"
        
        # For email, include the full traceback
        full_error_message = f"{error_message}\n\n{traceback.format_exc()}"
        
        print(full_error_message)
        # Send notifications
        send_notification_email("Birthday Cron Job: FAILED", full_error_message)
        send_telegram_notification(telegram_error_message)
        analytics = {'total_rows': 0, 'valid_emails': 0, 'birthdays_found': 0, 'emails_sent': 0, 'send_failures': 0, 'corrections': [], 'invalid_emails': [], 'skipped_sends': 0}  # Default on failure

    # Print analytics
    print("\n--- Analytics ---")
    for k, v in analytics.items():
        if k in ['corrections', 'invalid_emails']:
            print(f"{k.replace('_', ' ').title()}: {len(v)}")
        else:
            print(f"{k.replace('_', ' ').title()}: {v}")
