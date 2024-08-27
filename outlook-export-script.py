import win32com.client
import os
from datetime import datetime

def safe_get_attribute(obj, attr, default="N/A"):
    try:
        return getattr(obj, attr)
    except AttributeError:
        return default

def export_emails_to_text(output_folder):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 corresponds to the inbox

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    total_emails = inbox.Items.Count
    for i, message in enumerate(inbox.Items, 1):
        try:
            subject = safe_get_attribute(message, 'Subject')
            sender = safe_get_attribute(message, 'SenderName')
            date = safe_get_attribute(message, 'ReceivedTime')
            if isinstance(date, datetime):
                date = date.strftime("%Y-%m-%d %H:%M:%S")
            body = safe_get_attribute(message, 'Body')

            # Create a filename based on the date and subject
            filename = f"{i:04d}_{date.replace(':', '-')}_{subject[:30]}.txt"
            filename = "".join(c for c in filename if c.isalnum() or c in (' ', '.', '_', '-'))[:255]
            
            with open(os.path.join(output_folder, filename), 'w', encoding='utf-8') as f:
                f.write(f"From: {sender}\n")
                f.write(f"Date: {date}\n")
                f.write(f"Subject: {subject}\n\n")
                f.write(body)

            print(f"Exported {i}/{total_emails} emails", end='\r')

        except Exception as e:
            print(f"\nError exporting email {i}: {str(e)}")
            continue

if __name__ == "__main__":
    # Adjust this line
    output_folder =  r"C:\path\to\your\output"
    export_emails_to_text(output_folder)
    print("\nEmail export completed.")