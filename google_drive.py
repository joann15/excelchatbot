import os

from google.cloud import storage

BUCKET_NAME = "excelchatbot-storage"

SERVICE_ACCOUNT_FILE = "service-account.json"

client = storage.Client.from_service_account_json(
    SERVICE_ACCOUNT_FILE
)

bucket = client.bucket(BUCKET_NAME)


# ==========================
# Upload
# ==========================

def upload_file(filepath):

    blob_name = os.path.basename(filepath)

    blob = bucket.blob(blob_name)

    blob.upload_from_filename(filepath)

    return blob_name


# ==========================
# Download
# ==========================

def download_file(blob_name):

    local_path = f"temp_{blob_name}"

    blob = bucket.blob(blob_name)

    blob.download_to_filename(local_path)

    return local_path


# ==========================
# Update
# ==========================

def update_file(blob_name, filepath):

    blob = bucket.blob(blob_name)

    blob.upload_from_filename(filepath)