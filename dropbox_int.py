import dropbox
from dropbox.exceptions import AuthError, ApiError
import pandas as pd
import io
from logger import logs
from datetime import datetime
from dotenv import load_dotenv
import os

load_dotenv()


DROPBOX_ACCESS_TOKEN = os.getenv("DROPBOX_ACCESS_TOKEN")


def connect_to_dropbox(token: str):
    try:
        dbx = dropbox.Dropbox(token)
        dbx.users_get_current_account()
        print("✅ Connected to Dropbox.")
        logs.insert_one({"timestamp": datetime.utcnow(),"content": f"✅ Connected to Dropbox."})
        return dbx
    except AuthError as err:
        print(f"❌ Authentication error: {err}")
        logs.insert_one({"timestamp": datetime.utcnow(),"content": f"❌ Authentication error: {err}"})
        raise

def load_file_into_memory(dbx, file_path: str):
    try:
        _, res = dbx.files_download(path=file_path)
        return res.content
    except ApiError as err:
        print(f"❌ Error loading file {file_path}: {err}")
        logs.insert_one({"timestamp": datetime.utcnow(),"content": f"❌ Error loading file {file_path}: {err}"})
        raise

def list_files(dbx, folder_path: str = ""):
    try:
        result = dbx.files_list_folder(folder_path)
        files = []

        while True:
            for entry in result.entries:
                if isinstance(entry, dropbox.files.FileMetadata):
                    files.append(entry.path_display)

            if not result.has_more:
                break
            result = dbx.files_list_folder_continue(result.cursor)

        return files
    except ApiError as err:
        print(f"❌ API error: {err}")
        logs.insert_one({"timestamp": datetime.utcnow(),"content": f"❌ API error: {err}"})
        raise

def dropbox_connector():
    dbx = connect_to_dropbox(DROPBOX_ACCESS_TOKEN)
    folder_path = "/Lyzr test/penn stainless"

    files = list_files(dbx, folder_path)
    all_sheets_data = {}

    for file_path in files:
        if file_path.endswith(".xlsx"):
            content = load_file_into_memory(dbx, file_path)
            try:
                all_sheets = pd.read_excel(io.BytesIO(content), sheet_name=None)
                filename = file_path.split("/")[-1]

                for sheet_name, df in all_sheets.items():
                    key = f"{filename} - {sheet_name}"
                    all_sheets_data[key] = df
                    print(f"✅ Loaded: {key} with shape {df.shape}")
                    logs.insert_one({"timestamp": datetime.utcnow(),"content": f"✅ Loaded: {key} with shape {df.shape}"})
            except Exception as e:
                print(f"❌ Error reading Excel file {file_path}: {e}")
                logs.insert_one({"timestamp": datetime.utcnow(),"content": f"❌ Error reading Excel file {file_path}: {e}"})

    # Optional: print all sheet data
    for key, df in all_sheets_data.items():
        print(f"\n📄 {key}:\n")
        logs.insert_one({"timestamp": datetime.utcnow(),"content": f"\n📄 {key}:\n"})
        print(df.to_string(index=False))


    return all_sheets_data
