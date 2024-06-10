import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
from openpyxl import load_workbook

def send_email(smtp_server, port, sender_email, sender_password, recipient_email, subject, body):
    try:
        # Set up the server
        server = smtplib.SMTP(smtp_server, port)
        server.starttls()

        # Log in to the server
        server.login(sender_email, sender_password)

        # Create the email
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        # Send the email
        server.send_message(msg)
        server.quit()

        print("Email sent successfully")
    except Exception as e:
        print(f"Failed to send email: {e}")


# Example usage
# send_email('smtp.gmail.com', 587, 'your_email@gmail.com', 'your_password', 'recipient_email@gmail.com', 'Subject','Email body')


def rename_files(directory, old_pattern, new_pattern):
    try:
        for filename in os.listdir(directory):
            if old_pattern in filename:
                new_filename = filename.replace(old_pattern, new_pattern)
                old_file = os.path.join(directory, filename)
                new_file = os.path.join(directory, new_filename)
                os.rename(old_file, new_file)
        print("Files renamed successfully")
    except Exception as e:
        print(f"Failed to rename files: {e}")

# Example usage
# rename_files('/path/to/directory', 'old_pattern', 'new_pattern')

def update_spreadsheet(file_path, sheet_name, cell, value):
    try:
        workbook = load_workbook(filename=file_path)
        sheet = workbook[sheet_name]
        sheet[cell] = value
        workbook.save(file_path)
        print("Spreadsheet updated successfully")
    except Exception as e:
        print(f"Failed to update spreadsheet: {e}")

# Example usage
# update_spreadsheet('path/to/file.xlsx', 'Sheet1', 'A1', 'New Value')

def main():
    while True:
        print("\nTask Automation Menu:")
        print("1. Send Email")
        print("2. Rename Files")
        print("3. Update Spreadsheet")
        print("4. Exit")

        choice = input("Enter your choice: ")

        if choice == '1':
            smtp_server = input("Enter SMTP server: ")
            port = int(input("Enter port: "))
            sender_email = input("Enter sender email: ")
            sender_password = input("Enter sender password: ")
            recipient_email = input("Enter recipient email: ")
            subject = input("Enter subject: ")
            body = input("Enter email body: ")
            send_email(smtp_server, port, sender_email, sender_password, recipient_email, subject, body)

        elif choice == '2':
            directory = input("Enter directory path: ")
            old_pattern = input("Enter old pattern: ")
            new_pattern = input("Enter new pattern: ")
            rename_files(directory, old_pattern, new_pattern)

        elif choice == '3':
            file_path = input("Enter spreadsheet file path: ")
            sheet_name = input("Enter sheet name: ")
            cell = input("Enter cell to update (e.g., A1): ")
            value = input("Enter new value: ")
            update_spreadsheet(file_path, sheet_name, cell, value)

        elif choice == '4':
            break

        else:
            print("Invalid choice, please try again.")


if __name__ == "__main__":
    main()
