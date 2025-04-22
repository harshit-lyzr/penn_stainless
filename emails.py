from email.header import decode_header
from datetime import datetime
import pandas as pd
from io import BytesIO
from logger import logs


def clean_subject(subject):
    decoded = decode_header(subject)[0]
    if isinstance(decoded[0], bytes):
        return decoded[0].decode(decoded[1] if decoded[1] else "utf-8")
    return decoded[0]


def extract_email_body(msg):
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))
            if content_type == "text/plain" and "attachment" not in content_disposition:
                try:
                    return part.get_payload(decode=True).decode(errors="replace")
                except:
                    return ""
    else:
        return msg.get_payload(decode=True).decode(errors="replace")
    return ""


def process_excel_attachment(part, filename):
    file_data = part.get_payload(decode=True)
    excel_file = BytesIO(file_data)
    xls = pd.ExcelFile(excel_file)

    print(f"\n📄 Excel file: {filename}")
    logs.insert_one({"timestamp": datetime.utcnow(),"content": f"\n📄 Excel file: {filename}"})
    for sheet_name in xls.sheet_names:
        print(f"\n📑 Sheet: {sheet_name}")
        logs.insert_one({"timestamp": datetime.utcnow(),"content": f"\n📑 Sheet: {sheet_name}"})
        df = pd.read_excel(xls, sheet_name)
        print(df.to_string(index=False))
