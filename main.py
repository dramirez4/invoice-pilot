import json
import imaplib
import email
import os
import requests
import time
import datetime
from email.header import decode_header
from dotenv import load_dotenv

load_dotenv()

# Your email credentials
EMAIL = os.getenv("EMAIL")
PASSWORD = os.getenv("PASSWORD")
IMAP_SERVER = "imap.gmail.com"  # Example: 'imap.gmail.com' for Gmail

# PDF.co API key
API_KEY = os.getenv("API_KEY")

# Base URL for PDF.co Web API requests
BASE_URL = "https://api.pdf.co/v1"


def connect_email():
    """Connect to the email server and return the mail object."""
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL, PASSWORD)
    return mail


def search_invoices(mail):
    """Search for emails with the subject 'invoice'."""
    mail.select("inbox")
    status, messages = mail.search(None, '(SUBJECT "invoice")')
    email_ids = messages[0].split()
    return email_ids


def download_pdf_attachment(mail, email_id):
    """Download PDF attachment from the email."""
    status, data = mail.fetch(email_id, "(RFC822)")
    for response_part in data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])
            for part in msg.walk():
                if part.get_content_type() == "application/pdf":
                    filename = part.get_filename()
                    if filename:
                        filename = decode_header(filename)[0][0]
                        if isinstance(filename, bytes):
                            filename = filename.decode()
                        with open(filename, "wb") as f:
                            f.write(part.get_payload(decode=True))
                            return filename
    return None


def get_parsed_invoice(file_path):
    """Send the invoice PDF to PDF.co for processing and return the parsed data."""
    with open(file_path, "rb") as file:
        # Upload the file to PDF.co and get the file URL
        url = f"{BASE_URL}/file/upload/get-presigned-url?name={os.path.basename(file_path)}"
        headers = {"x-api-key": API_KEY}
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            json_response = response.json()
            if json_response["error"] == False:
                upload_url = json_response["presignedUrl"]
                file_url = json_response["url"]
                
                # Upload the actual file to the presigned URL
                with open(file_path, "rb") as file:
                    upload_response = requests.put(upload_url, data=file)
                    if upload_response.status_code == 200:
                        return parse_invoice(file_url)
            else:
                print("Error uploading file:", json_response["message"])
    return None


def parse_invoice(uploaded_file_url):
    """Send the invoice to PDF.co AI Invoice Parser API."""
    parameters = {"url": uploaded_file_url}
    url = f"{BASE_URL}/ai-invoice-parser"
    headers = {"x-api-key": API_KEY}
    response = requests.post(url, data=parameters, headers=headers)
    if response.status_code == 200:
        json_response = response.json()
        if json_response["error"] == False:
            job_id = json_response["jobId"]
            return check_job_status(job_id)
        else:
            print("Error in parsing invoice:", json_response["message"])
    else:
        print("Error with request:", response.status_code, response.reason)
    return None


def check_job_status(job_id):
    """Check the job status until the job is finished."""
    while True:
        url = f"{BASE_URL}/job/check?jobid={job_id}"
        headers = {"x-api-key": API_KEY}
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            json_response = response.json()
            status = json_response["status"]
            print(f"Job status: {status}")
            if status == "success":
                # The job has finished, return the result
                return json_response
            elif status == "working":
                time.sleep(3)  # Wait a bit before checking again
            else:
                print(f"Job failed: {status}")
                break
        else:
            print("Error with job status request:", response.status_code, response.reason)
            break
    return None


def save_json(data, file_name):
    """Save parsed data to a JSON file."""
    with open(file_name, "w") as json_file:
        json.dump(data, json_file, indent=4)
    print(f"Saved parsed invoice to {file_name}")


def main():
    # Step 1: Connect to email
    mail = connect_email()

    # Step 2: Search for invoice emails
    email_ids = search_invoices(mail)
    if not email_ids:
        print("No invoices found.")
        return

    # Step 3: Process each invoice email
    for email_id in email_ids:
        # Step 4: Download the PDF attachment from the email
        pdf_filename = download_pdf_attachment(mail, email_id)
        if pdf_filename:
            print(f"Downloaded PDF: {pdf_filename}")

            # Step 5: Use PDF.co API to process the invoice
            parsed_data = get_parsed_invoice(pdf_filename)
            if parsed_data:
                json_filename = pdf_filename.replace(".pdf", ".json")
                # Step 6: Save the parsed invoice data to a JSON file
                save_json(parsed_data, json_filename)
            else:
                print(f"Failed to parse invoice: {pdf_filename}")
        else:
            print(f"No PDF found in email {email_id}")


if __name__ == "__main__":
    main()







