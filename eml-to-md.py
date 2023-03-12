import os
import email
import markdown
import shutil
from datetime import datetime
from dateutil import parser
from dateutil import tz

# Folder path containing .eml files
folder_path = '/Users/knight/Desktop/Imported/Sent Items.mbox/eml'

# Create attachments folder if it doesn't exist
attachments_folder = os.path.join(folder_path, 'attachments')
if not os.path.exists(attachments_folder):
    os.makedirs(attachments_folder)

# Create 'done' folder if it doesn't exist
done_folder = os.path.join(folder_path, 'done')
if not os.path.exists(done_folder):
    os.makedirs(done_folder)

# Function to remove illegal characters from filename
def clean_filename(filename):
    cleaned = ''.join(c if c.isalnum() or c in [' ', '.', '_'] else '_' for c in filename)
    return cleaned.strip()

# Loop through all .eml files in the folder
for filename in os.listdir(folder_path):
    if filename.endswith('.eml'):
        try:
            # Parse email message
            with open(os.path.join(folder_path, filename), 'rb') as f:
                try:
                    # Try decoding with UTF-8 first
                    contents = f.read().decode('utf-8')
                except UnicodeDecodeError:
                    # Fall back to binary decoding
                    f.seek(0)
                    contents = f.read()
                if isinstance(contents, bytes):
                    msg = email.message_from_bytes(contents)
                else:
                    msg = email.message_from_string(contents)

            # Extract date, sender, and subject from email message
            date_str = msg['Date']
            sender = msg['From']
            subject = msg['Subject']

            # Convert email body to markdown
            body = ''
            for part in msg.walk():
                if part.get_content_type() == 'text/plain':
                    body += part.get_payload()
                elif part.get_content_maintype() == 'multipart':
                    continue
                else:
                    # Download attachment and create link
                    filename = part.get_filename()
                    if filename:
                        filepath = os.path.join(attachments_folder, filename)
                        with open(filepath, 'wb') as f:
                            f.write(part.get_payload(decode=True))
                        link = f"[{filename}]({filename})"
                        body += f"\n\n{link}"

            body = markdown.markdown(body)

            # Parse date string into datetime object
            try:
                date = parser.parse(date_str)
                date = date.astimezone(tz.tzlocal()).replace(tzinfo=None)
            except (ValueError, TypeError):
                with open('error.log', 'a') as logfile:
                    logfile.write(f"Error parsing date for file {filename}\n")
                continue

            # Format date as (YYYY-MM-DD) HH:MM:SS
            date_str = f"({date.strftime('%Y-%m-%d')}) {date.strftime('%H-%M-%S')}"

            # Clean up filename
            sender_cleaned = clean_filename(sender)
            subject_cleaned = clean_filename(subject)
            new_filename = f"{date_str} - {sender_cleaned} - {subject_cleaned}.md"

            # Write markdown file
            with open(os.path.join(folder_path, new_filename), 'w') as f:
                f.write(body)

            # Move processed .eml file to 'done' folder
            shutil.move(os.path.join(folder_path, filename), os.path.join(done_folder, filename))

        except Exception as e:
            with open('error.log', 'a') as logfile:
                logfile.write(f"Error processing file {filename}: {e}\n")