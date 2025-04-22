import imaplib
import email
import os
from pydantic import BaseModel,Field
from emails import extract_email_body, clean_subject
from dropbox_int import dropbox_connector
from lyzr_agent import send_message_to_agent
import json
from utils import create_excel_file,send_email_with_attachment
from datetime import datetime
from logger import logs
from dotenv import load_dotenv

load_dotenv()


# Email credentials and server
EMAIL = os.getenv("EMAIL")
PASSWORD = os.getenv("PASSWORD")
IMAP_SERVER = os.getenv("IMAP_SERVER")


class Logs(BaseModel):
    timestamp: datetime = Field(default_factory=datetime.utcnow)
    content: str


def check_unseen_emails():
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL, PASSWORD)
        logs.insert_one({"content": f"Login Successfull {EMAIL}"})
        mail.select("inbox")

        result, data = mail.search(None, 'UNSEEN')
        if result != "OK":
            print("❌ Failed to search for unseen emails.")
            logs.insert_one({"timestamp": datetime.utcnow(),"content": f"❌ Failed to search for unseen emails."})
            return

        email_ids = data[0].split()
        print(f"📬 You have {len(email_ids)} unseen emails.\n")
        logs.insert_one({"timestamp": datetime.utcnow(),"content": f"📬 You have {len(email_ids)} unseen emails.\n"})

        for eid in email_ids:
            result, msg_data = mail.fetch(eid, "(RFC822)")
            if result != "OK":
                print(f"⚠️ Could not fetch email ID {eid}")
                logs.insert_one({"timestamp": datetime.utcnow(),"content": f"⚠️ Could not fetch email ID {eid}"})
                continue

            msg = email.message_from_bytes(msg_data[0][1])
            subject = clean_subject(msg.get("Subject", "No Subject"))
            from_ = msg.get("From", "Unknown Sender")
            body = extract_email_body(msg)

            print(f"From: {from_}")
            logs.insert_one({"timestamp": datetime.utcnow(),"content": f"From: {from_}"})
            print(f"Subject: {subject}")
            logs.insert_one({"timestamp": datetime.utcnow(),"content": f"Subject: {subject}"})
            print(f"Body:\n{body}\n")
            logs.insert_one({"timestamp": datetime.utcnow(),"content": f"Body:\n{body}\n"})

            # Process attachments
            # for part in msg.walk():
            #     content_disposition = str(part.get("Content-Disposition"))
            #     if part.get_content_maintype() == "multipart":
            #         continue
            #     if "attachment" in content_disposition:
            #         filename = part.get_filename()
            #         if filename:
            #             filename = decode_header(filename)[0][0]
            #             if isinstance(filename, bytes):
            #                 filename = filename.decode()
            #             if filename.endswith((".xlsx", ".xls")):
            #                 process_excel_attachment(part, filename)

            print("-" * 80)
            mail.store(eid, '+FLAGS', '\\Seen')

            dropbox_data = dropbox_connector()
            # Example usage
            response = send_message_to_agent(
                user_id="harshit@lyzr.ai",
                agent_id="6805360623b40bb3326101e2",
                message=f"{dropbox_data} Quote Request: {body}",
                api_key=os.getenv("LYZR_KEY")
            )
            print(response)

            da = json.loads(response['response'])
            print(da)

            filename = "outputs.xlsx"
            customer_name = "Harshit"
            date = "03/05/2025"
            address = "Surat"

            create_excel_file(filename, da['RFQ_Details'], customer_name, date, address)
            send_email_with_attachment(
                subject=subject,
                body="Hello, \n\nPlease find attached the quotation as per your request. The details are outlined in the attached Excel file for your review.",
                sender_email=EMAIL,
                sender_password=PASSWORD,
                receiver_email=from_,
                attachment_path=filename
            )



        mail.logout()

    except imaplib.IMAP4.error as e:
        print(f"IMAP error: {e}")
    except Exception as e:
        print(f"Unexpected error: {e}")


if __name__ == "__main__":
    check_unseen_emails()