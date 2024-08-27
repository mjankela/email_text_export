# Email exporter script to plain text
A set of scripts that allow you to export your email content into plain text.

# outlook-export-script.py
This script uses the win32com library to interact with Outlook and export all emails from the inbox to individual text files. 
Each file contains the sender, date, subject, and body of the email.

To use these scripts:

Install the required library by running "pip install pywin32" in your command prompt or terminal.

Adjust line 46  output_folder =  r"C:\path\to\your\output"

    # Adjust this line
    output_folder =  r"C:\path\to\your\output"

Run the  script to export all emails to individual text files with python .\outlook-export-script.py

# merge-emails-script.py
This script reads all the text files in the specified folder and merges them into a single large text file, separating each email with a delimiter.
The maximum size of the text file can be configured. Default is 4MB.

Adjust jere line 6: def merge_text_files(input_folder, output_base_path, max_size_bytes=4 * 1024 * 1024):  # 4 MB

To use these scripts adjust line 36 and 37:

    # Adjust this line
    input_folder = r"C:\path\to\your\input\folder"
    output_base_path = r"C:\path\to\your\merged_emails"

then run python .\merge-emails-script.py
