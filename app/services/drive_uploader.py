import json
import logging
import os

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

logger = logging.getLogger(__name__)

SCOPES = ['https://www.googleapis.com/auth/drive']
DRIVE_FOLDER_ID = os.getenv('DRIVE_FOLDER_ID', '')

DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'


def _get_drive_service():
    raw = os.getenv('GOOGLE_SERVICE_ACCOUNT_JSON')
    if not raw:
        raise EnvironmentError('GOOGLE_SERVICE_ACCOUNT_JSON env var is not set')
    info = json.loads(raw)
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return build('drive', 'v3', credentials=creds)


def upload_file(local_path: str, filename: str, folder_id: str = '') -> dict:
    """
    Upload a file to Google Drive inside folder_id.
    Returns {'file_id', 'view_url', 'download_url'}.
    Raises on failure so the caller can decide to fall back.
    """
    target_folder = folder_id or DRIVE_FOLDER_ID
    if not target_folder:
        raise ValueError('DRIVE_FOLDER_ID is not set — cannot upload to Drive')

    service = _get_drive_service()

    file_metadata = {'name': filename, 'parents': [target_folder]}
    media = MediaFileUpload(local_path, mimetype=DOCX_MIME, resumable=False)

    created = service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id,webViewLink,webContentLink',
        supportsAllDrives=True,
    ).execute()

    file_id = created['id']

    # Make readable by anyone with the link so GAS can open it in a browser
    service.permissions().create(
        fileId=file_id,
        body={'type': 'anyone', 'role': 'reader'},
        supportsAllDrives=True,
    ).execute()

    logger.info(f'Uploaded {filename} → Drive file_id={file_id}')

    return {
        'file_id': file_id,
        'view_url': created.get('webViewLink', ''),
        'download_url': created.get('webContentLink', ''),
    }
