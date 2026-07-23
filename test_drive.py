from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import os

SCOPES = ["https://www.googleapis.com/auth/drive"]

credentials = service_account.Credentials.from_service_account_file(
    "service-account.json",
    scopes=SCOPES
)

service = build("drive", "v3", credentials=credentials)

FOLDER_ID = "1gFQsUKSOc-9wD0LVGWYwi6R9x7tqfLIW"

file_metadata = {
    "name": "hello.txt",
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

print(file)