import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Load Excel file without headers
def load_excel(file_path):
    """
    Load the Excel file without headers.

    Args:
        file_path (str): The file path of the Excel file.

    Returns:
        DataFrame: A pandas DataFrame containing the Excel data.
    """
    df = pd.read_excel(file_path, header=None)  # Reading Excel file without headers
    return df

# Filter emails based on a condition (if needed)
def filter_emails(df, column_name, condition):
    """
    Filter rows from the DataFrame based on a condition.

    Args:
        df (DataFrame): The input DataFrame.
        column_name (int): The index of the column to apply the condition.
        condition (str): The condition to filter rows by.

    Returns:
        DataFrame: Filtered DataFrame based on the given condition.
    """
    selected_rows = df[df[column_name] == condition]  # Filter rows based on condition
    return selected_rows

# Send email to a recipient
def send_email(sender_email, sender_password, recipient_email, subject, body):
    """
    Send an email to a single recipient.

    Args:
        sender_email (str): The sender's email address.
        sender_password (str): The sender's email password.
        recipient_email (str): The recipient's email address.
        subject (str): The email subject.
        body (str): The body of the email.

    Returns:
        None
    """
    # Create the MIME message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject

    # Attach the email body
    msg.attach(MIMEText(body, 'plain'))

    # Set up the SMTP server (Gmail SMTP server is used here)
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()

    try:
        # Log in to the sender's email account
        server.login(sender_email, sender_password)

        # Send the email
        server.sendmail(sender_email, recipient_email, msg.as_string())
        print(f"Email sent to {recipient_email}")
    except Exception as e:
        # If there is an issue, print the error
        print(f"Failed to send email to {recipient_email}. Error: {str(e)}")
    finally:
        # Close the SMTP connection
        server.quit()

# Main function to execute the mass emailing process
def mass_mailer(file_path, sender_email, sender_password, subject, body, email_column_index):
    """
    Send bulk emails by reading email addresses from an Excel file.

    Args:
        file_path (str): Path to the Excel file containing email addresses.
        sender_email (str): The sender's email address.
        sender_password (str): The sender's email password.
        subject (str): The email subject.
        body (str): The email body.
        email_column_index (int): Index of the email column in the Excel file.

    Returns:
        None
    """
    # Load the Excel file
    df = load_excel(file_path)

    # Iterate over each row and send email to each recipient
    for index, row in df.iterrows():
        recipient_email = row[email_column_index]  # Extract email from the specified column
        send_email(sender_email, sender_password, recipient_email, subject, body)

# Example usage
if __name__ == "__main__":
    # Path to the Excel file
    excel_file = "mail.xlsx"  # Update this with the path to your Excel file
    
    # Email account credentials (Update with actual credentials)
    sender_email = "your_email@example.com"
    sender_password = "your_password"
    
    # Email subject and body
    subject = "Reminder"
    body = "Hi, how are you?"
    
    # Index for the column containing email addresses (0-indexed, so 2 corresponds to Column C)
    email_column_index = 2

    # Run the mass mailer function
    mass_mailer(excel_file, sender_email, sender_password, subject, body, email_column_index)
