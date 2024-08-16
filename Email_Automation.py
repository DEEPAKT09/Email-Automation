import os
import requests
import sys
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from urllib.parse import urlparse
import argparse

# Function to extract spreadsheet ID from the Google Sheets URL
def extract_spreadsheet_id(url):
    parsed_url = urlparse(url)
    spreadsheet_id = parsed_url.path.split('/')[3]
    return spreadsheet_id

# Function to download Google Sheet as a TSV file
def getGoogleSheet(spreadsheet_id, outDir, outFile):
    url = f'https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=tsv'
    response = requests.get(url)
    if response.status_code == 200:
        filepath = os.path.join(outDir, outFile)
        with open(filepath, 'wb') as f:
            f.write(response.content)
            print('TSV file saved to: {}'.format(filepath))
        return filepath  # Return the file path
    else:
        print(f'Error downloading Google Sheet: {response.status_code}')
        sys.exit(1)

# Function to extract emails, subjects, and body messages from the TSV file
def extract_email_data_from_tsv(filepath):
    email_data = []
    if not os.path.exists(filepath):
        print("File does not exist.")
        return email_data
    
    with open(filepath, 'r') as f:
        reader = pd.read_csv(f, delimiter='\t')
        
        # Extract relevant columns for each row and convert to string
        for _, row in reader.iterrows():
            email = str(row.get('Email', '')).strip()
            subject = str(row.get('Subject', '')).strip()
            body = str(row.get('Body', '')).strip()

            # Only add rows with non-empty email addresses
            if email:
                email_data.append({
                    'Email': email,
                    'Subject': subject,
                    'Body': body
                })

    return email_data  # Return the list of dictionaries with email data

# Function to send bulk emails
def send_bulk_emails(email_data, sender_email, sender_password):
    smtp_server = "smtp.gmail.com"
    smtp_port = 587

    try:
        # Establish a connection to the SMTP server
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)

        for data in email_data:
            recipient = data['Email']
            subject = data['Subject']
            body = data['Body']

            # Create a MIME message object
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = recipient
            msg['Subject'] = subject

            # Attach the body to the message
            msg.attach(MIMEText(body, 'plain'))

            # Send the email
            server.sendmail(sender_email, recipient, msg.as_string())

        # Close the connection to the server
        server.quit()
        print("Emails sent successfully!")

    except Exception as e:
        print(f"Error: {str(e)}")

# Main execution flow
if __name__ == "__main__":
    # Set up argument parser
    parser = argparse.ArgumentParser(description="Send bulk emails based on Google Sheets data.")
    parser.add_argument('-i', '--input', required=True, help="Google Sheets URL")

    # Parse arguments
    args = parser.parse_args()
    google_sheets_url = args.input

    # Extract spreadsheet ID from the URL
    spreadsheet_id = extract_spreadsheet_id(google_sheets_url)

    # Set output directory and ensure it exists
    outDir = 'Output Directory Path Here'
    os.makedirs(outDir, exist_ok=True)

    # Download the Google Sheet and save it as a TSV
    filepath = getGoogleSheet(spreadsheet_id, outDir, "SAMPLE.tsv")

    # Extract email data (email, subject, body) from the downloaded TSV file
    email_data = extract_email_data_from_tsv(filepath)

    # Credentials for sending emails
    sender_email = "Email Address of the sender"
    sender_password = "Password of sender email address"

    # Send bulk emails using the extracted email data
    send_bulk_emails(email_data, sender_email, sender_password)

    sys.exit(0)  # Success
