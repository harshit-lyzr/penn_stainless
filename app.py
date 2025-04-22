import imaplib
import email
import os
import json
import time
from pydantic import BaseModel, Field
from emails import extract_email_body, clean_subject
from dropbox_int import dropbox_connector
from lyzr_agent import send_message_to_agent
from utils import create_excel_file, send_email_with_attachment
from datetime import datetime
from logger import logs
from dotenv import load_dotenv
from concurrent.futures import ThreadPoolExecutor

load_dotenv()

EMAIL = os.getenv("EMAIL")
PASSWORD = os.getenv("PASSWORD")
IMAP_SERVER = os.getenv("IMAP_SERVER")


class Logs(BaseModel):
    timestamp: datetime = Field(default_factory=datetime.utcnow)
    content: str


def process_email(mail, eid):
    start_time = time.time()
    log_buffer = []

    try:
        result, msg_data = mail.fetch(eid, "(RFC822)")
        if result != "OK":
            log_buffer.append({"timestamp": datetime.utcnow(), "content": f"⚠️ Could not fetch email ID {eid}"})
            return

        msg = email.message_from_bytes(msg_data[0][1])
        subject = clean_subject(msg.get("Subject", "No Subject"))
        from_ = msg.get("From", "Unknown Sender")
        body = extract_email_body(msg)

        log_buffer.append({"timestamp": datetime.utcnow(), "content": f"From: {from_}"})
        log_buffer.append({"timestamp": datetime.utcnow(), "content": f"Subject: {subject}"})
        log_buffer.append({"timestamp": datetime.utcnow(), "content": f"Body:\n{body}"})


        dropbox_data = dropbox_connector()

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

        create_excel_file(filename, da[0]['RFQ_Details'], customer_name, date, address)

        send_email_with_attachment(
            subject=subject,
            body=da[0]['email_body'],
            sender_email=EMAIL,
            sender_password=PASSWORD,
            receiver_email=from_,
            attachment_path=filename
        )
        # mail.store(eid, '+FLAGS', '\\Seen')
        log_buffer.append({"timestamp": datetime.utcnow(), "content": f"✅ Processed email from {from_} in {time.time() - start_time:.2f}s"})

    except Exception as e:
        log_buffer.append({"timestamp": datetime.utcnow(), "content": f"❌ Error processing email {eid}: {str(e)}"})

    finally:
        if log_buffer:
            logs.insert_many(log_buffer)


def check_unseen_emails():
    overall_start = time.time()
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL, PASSWORD)
        logs.insert_one({"content": f"✅ Login successful for {EMAIL}"})
        mail.select("inbox")

        result, data = mail.search(None, 'UNSEEN')
        if result != "OK":
            logs.insert_one({"timestamp": datetime.utcnow(), "content": "❌ Failed to search for unseen emails."})
            return

        email_ids = data[0].split()

        if not email_ids:
            print("No unseen emails.")
            return
        print(f"📬 You have {len(email_ids)} unseen emails.")
        logs.insert_one({"timestamp": datetime.utcnow(), "content": f"📬 You have {len(email_ids)} unseen emails."})

        with ThreadPoolExecutor(max_workers=5) as executor:
            for eid in email_ids:
                executor.submit(process_email, mail, eid)

        mail.logout()
        print(f"✅ Completed processing {len(email_ids)} emails in {time.time() - overall_start:.2f} seconds")

    except imaplib.IMAP4.error as e:
        print(f"IMAP error: {e}")
    except Exception as e:
        print(f"Unexpected error: {e}")


if __name__ == "__main__":
    check_unseen_emails()
