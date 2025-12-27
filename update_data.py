import shutil
import os
from dotenv import load_dotenv

load_dotenv()
DATA_SOURCE = os.getenv('DATA_SOURCE', '/path/to/new/December.xlsx')  # Path to new file (e.g., uploaded via SCP)

def update_excel():
    if os.path.exists(DATA_SOURCE):
        shutil.copy(DATA_SOURCE, 'December.xlsx')
        print("Excel file updated successfully.")
        # Optionally, clear caches to force re-validation
        if os.path.exists('email_cache.pkl'):
            os.remove('email_cache.pkl')
        if os.path.exists('sent_log.pkl'):
            os.remove('sent_log.pkl')
    else:
        print("New data file not found.")

if __name__ == "__main__":
    update_excel()