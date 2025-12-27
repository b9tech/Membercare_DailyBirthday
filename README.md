# Daily Birthday Celebration Cronjob

This project is a daily cron job to automatically send birthday wishes to members of the NCS. It reads a list of members from an Excel file, checks for birthdays on the current day, and sends individual emails with a birthday image to each celebrant for privacy.

## Setup

### 1. Install Dependencies

First, install the required Python libraries using pip. Make sure you are in the project directory in your terminal.

```bash
pip install -r requirements.txt
```

### 2. Configure Environment Variables

Create a file named `.env` in the project root directory. This file will store your sensitive email credentials. Add the following lines to the `.env` file, replacing the placeholder values with your actual information.

```
# --- Email Sender Credentials ---
EMAIL_SENDER=your_email@example.com
EMAIL_PASSWORD=your_app_password
SMTP_SERVER=mail.ncs.org.ng
SMTP_PORT=587

# --- Notification Settings ---
# Your email addresses for status notifications (comma-separated)
ADMIN_EMAILS=your_admin_email1@example.com, your_admin_email2@example.com

# (Optional) Telegram Bot credentials for status notifications
TELEGRAM_BOT_TOKEN=
TELEGRAM_CHAT_ID=
```

### 3. Prepare the Member List

Place your Excel file named `December.xlsx` in the same directory as the script. The file must have at least two columns:
- `EMAIL`: The email address of the member.
- `DOB`: The member's birthday in `YYYY-MM-DD` format.

Here is an example of how the Excel file should be structured:

| Name      | EMAIL                  | DOB        |
|-----------|------------------------|------------|
| John Doe  | john.doe@example.com   | 1990-12-21 |
| Jane Smith| jane.smith@example.com | 1985-04-15 |

### 4. Add the Birthday Image

Place the birthday image you want to send in the same directory as the script and name it `birthday.png`.

## Autonomous Notifications

To provide true "fire-and-forget" autonomy, the script now includes a notification system that can alert you via Email and Telegram.

### Email Notifications
After every run, the script will send a status email to the `ADMIN_EMAIL` you configured in your `.env` file.
- **On Success:** You will receive an email with the subject `Birthday Cron Job: Success`.
- **On Failure:** You will immediately receive an email with the subject `Birthday Cron Job: FAILED` containing the full error details.

### Telegram Notifications (Optional)
You can also receive instant notifications on your phone via Telegram.

**How to Set Up Telegram Notifications:**

1.  **Create a Telegram Bot:**
    *   Open Telegram and search for the `@BotFather` bot.
    *   Start a chat with BotFather and send the `/newbot` command.
    *   Follow the prompts to choose a name and username for your bot.
    *   BotFather will give you a unique **Bot Token**. Copy this token.

2.  **Get Your Chat ID:**
    *   Search for the `@userinfobot` bot on Telegram and start a chat with it.
    *   It will immediately reply with your user information, including your **Chat ID**. Copy this ID.

3.  **Update Your `.env` File:**
    *   Paste the **Bot Token** and **Chat ID** into the corresponding variables in your `.env` file.
    *   Make sure your bot can message you by sending it a `/start` command or any message first.

## Running the Script

You can run the script manually to test it or to send wishes for the current day.

```bash
python birthday_mailer.py
```

## Scheduling the Daily Cron Job on Windows

To run the script automatically every day, you can use the Windows Task Scheduler.

1.  **Open Task Scheduler:** Press `Win + R`, type `taskschd.msc`, and press Enter.
2.  **Create a New Task:** In the right-hand pane, click "Create Basic Task...".
3.  **Name the Task:** Give your task a name (e.g., "Daily Birthday Email") and a description, then click "Next".
4.  **Set the Trigger:** Select "Daily" and click "Next". Choose a time to run the script each day (e.g., 9:00 AM) and click "Next".
5.  **Set the Action:** Select "Start a program" and click "Next".
6.  **Configure the Action:**
    *   In the "Program/script" field, browse to the `python.exe` file inside your project's `.venv\Scripts` folder. The path will be something like `C:\Users\YourUser\AgenticAI\CronDailyBirthday\.venv\Scripts\python.exe`.
    *   In the "Add arguments (optional)" field, enter the name of your script: `birthday_mailer.py`.
    *   In the "Start in (optional)" field, enter the full path to your project directory: `C:\Users\YourUser\AgenticAI\CronDailyBirthday`.
7.  **Finish:** Review your settings and click "Finish".

The task is now scheduled to run at the time you specified. You can view and manage it from the Task Scheduler library.
