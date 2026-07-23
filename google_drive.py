import os
import io

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

SCOPES = ["https://www.googleapis.com/auth/drive"]

FOLDER_ID = "1gFQsUKSOc-9wD0LVGWYwi6R9x7tqfLIW"

SERVICE_ACCOUNT_FILE = "service_account.json"

creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE,
    scopes=SCOPES
)

service = build("drive", "v3", credentials=creds)


# ==========================
# Upload a new file
# ==========================
def upload_file(filepath):

    file_metadata = {
        "name": os.path.basename(filepath),
        "parents": [FOLDER_ID]
    }

    media = MediaFileUpload(
        filepath,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=False
    )

    file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id"
    ).execute()

    return file["id"]


# ==========================
# Download an existing file
# ==========================
def download_file(file_id):

    request = service.files().get_media(fileId=file_id)

    local_path = f"temp_{file_id}.xlsx"

    with open(local_path, "wb") as f:

        downloader = MediaIoBaseDownload(f, request)

        done = False

        while not done:
            status, done = downloader.next_chunk()

    return local_path


# ==========================
# Update an existing file
# ==========================
def update_file(file_id, filepath):

    media = MediaFileUpload(
        filepath,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=False
    )

    service.files().update(
        fileId=file_id,
        media_body=media
    ).execute()