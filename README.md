# Python Mass Emailer

This Python script allows you to send mass emails using an Excel sheet containing email addresses. The email addresses are read from a specified column in the Excel file, and emails are sent using the Gmail SMTP server.

## Features

- Reads email addresses from an Excel file without headers.
- Sends personalized emails to recipients using the Gmail SMTP server.
- Allows subject and body customization.
- Optionally, you can filter email addresses based on specific conditions.

## Requirements

- Python 3.x
- Pandas (for reading Excel files)
- smtplib (for sending emails)
- email (for formatting the email)

## Installation

1. Clone this repository to your local machine:

   ```bash
   https://github.com/Anshkumar111/MailBot.git

2. Install the required Python libraries:

   ```bash
   pip install pandas

3. Set up your Gmail account to allow access for less secure apps or generate an App Password if 2-factor authentication is enabled.

## Usage
1. Update the script with the path to your Excel file.

2. Replace the sender's email and password with your own Gmail credentials.

3. Set the desired email subject and body.

4. Specify the column index containing email addresses (zero-indexed).

5. Run the script:

   ```bash
   python mass_mailer.py
## Example
Assume you have an Excel file mail.xlsx with email addresses in column C. The script will read the emails from this column and send a reminder email to each recipient.

   ```python
   sender_email = "your_email@example.com"
   sender_password = "your_password"
   subject = "Reminder"
   body = "Hi, how are you?"
   email_column_index = 2  # Column C (zero-indexed)

