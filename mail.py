import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Load Excel file without headers
def load_excel(file_path):
    # Read Excel file without headers
    df = pd.read_excel(file_path, header=None)
    return df

# Filter emails based on a condition (if any)
def filter_emails(df, column_name, condition):
    # Example: Select all rows where the condition is True (if filtering is needed)
    selected_rows = df[df[column_name] == condition]
    return selected_rows

# Send email to the recipients
def send_email(sender_email, sender_password, recipient_email, subject, body):
    # Create the MIME message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject

    # Attach the email body
    msg.attach(MIMEText(body, 'plain'))

    # Set up the SMTP server (Gmail SMTP server in this example)
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()

    try:
        # Log in to the email account
        server.login(sender_email, sender_password)

        # Send the email
        server.sendmail(sender_email, recipient_email, msg.as_string())
        print(f"Email sent to {recipient_email}")
    except Exception as e:
        print(f"Failed to send email to {recipient_email}. Error: {str(e)}")
    finally:
        server.quit()

# Main function to execute the mass mailing
def mass_mailer(file_path, sender_email, sender_password, subject, body, email_column_index):
    # Load the Excel file
    df = load_excel(file_path)

    # Send emails
    for index, row in df.iterrows():
        recipient_email = row[email_column_index]  # Use the index of Column 'C', which is 2 (since Python is zero-indexed)
        send_email(sender_email, sender_password, recipient_email, subject, body)

# Example usage
if __name__ == "__main__":
    # Path to the Excel file
    excel_file = "mail.xlsx"  # Update with the path to your Excel file
    
    # Email account credentials
    sender_email = "Sender's_mail_ID" # Update with the sender mail-ID
    sender_password = "Sender's_mail_password-" #Update with the Sender mail-ID password
    
    # Email subject and body or message or information you want to send
    subject = "Reminder"
    body = "Hi, how are you?"

    # Index for column C (zero-indexed, so 2 corresponds to Column C)
    email_column_index = 2  # This corresponds to the third column (C)

    # Run the mass mailer
    mass_mailer(excel_file, sender_email, sender_password, subject, body, email_column_index)
