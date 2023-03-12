import os
import re
import hashlib
from email import message_from_file
from datetime import datetime

# Set directory paths
email_directory = "/Users/knight/Desktop/Imported/Tony @ Entelechy Design.mbox/eml"

# Create directories if they don't exist
attachments_directory = os.path.join(email_directory, "Attachments")
os.makedirs(attachments_directory, exist_ok=True)

markdown_directory = os.path.join(email_directory, "MD")
os.makedirs(markdown_directory, exist_ok=True)

eml_directory = os.path.join(email_directory, "EML")
os.makedirs(eml_directory, exist_ok=True)

# Open the error log file for writing
with open(os.path.join(email_directory, "Errors.txt"), "w") as error_file:
    for filename in os.listdir(email_directory):
        if not filename.endswith(".eml"):
            continue

        try:
            with open(os.path.join(email_directory, filename), "r", encoding="utf-8") as email_file:
                msg = message_from_file(email_file)

            # Extract header information
            email_from = msg["From"]
            email_subject = msg["Subject"]
            email_date = datetime.strptime(msg["Date"], "%a, %d %b %Y %H:%M:%S %z")

            # Generate strings for file name and content
            date_string = f"({email_date.strftime('%Y-%m-%d')}) {email_date.strftime('%H-%M-%S')}"
            sender_string = re.sub("[^0-9a-zA-Z]+", "-", email_from)
            subject_string = re.sub("[^0-9a-zA-Z]+", "-", email_subject)

            # Generate markdown content
            markdown_content = f"# {email_subject}\n\n"
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    markdown_content += part.get_payload()
                elif part.get_content_type().startswith("image/"):
                    # Save attachments to Attachments directory and update markdown content
                    attachment_name = part.get_filename()
                    attachment_path = os.path.join(attachments_directory, attachment_name)

                    # If an attachment with the same name exists, check for equality
                    if os.path.isfile(attachment_path):
                        with open(attachment_path, "rb") as existing_file:
                            existing_file_content = existing_file.read()
                            existing_file_hash = hashlib.sha256(existing_file_content).hexdigest()

                        part_hash = hashlib.sha256(part.get_payload(decode=True)).hexdigest()

                        if existing_file_hash == part_hash:
                            markdown_content += f"\n\n[Attachment: {attachment_name}]({attachment_path})"
                            continue
                        else:
                            name, extension = os.path.splitext(attachment_name)
                            attachment_name = f"{name}-1{extension}"

                    with open(attachment_path, "wb") as attachment_file:
                        attachment_file.write(part.get_payload(decode=True))

                    markdown_content += f"\n\n[Attachment: {attachment_name}]({os.path.relpath(attachment_path, markdown_directory)})"

            # Create markdown file
            markdown_filename = f"{date_string} - [{sender_string}] - [{subject_string}].md"
            with open(os.path.join(markdown_directory, markdown_filename), "w", encoding="utf-8") as markdown_file:
                markdown_file.write(markdown_content)

            # Move eml file to EML directory
            os.rename(os.path.join(email_directory, filename), os.path.join(eml_directory, filename))
        except Exception as e:

            # Log error and continue with next file
            error_file.write(f"{filename}: {e}\n")
        
