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

Install the required Python libraries:

   ```bash
   pip install pandas
Set up your Gmail account to allow access for less secure apps or generate an App Password if 2-factor authentication is enabled.
