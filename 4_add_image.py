import google.auth
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# --- CONFIGURATION ---
SERVICE_ACCOUNT_FILE = 'service_account.json'
EXISTING_DOC_ID = '10YTudX4LmIjxsoT2gOJFw2-yKMRYCxSJTNkHf4hZqoE'
IMAGE_FILENAME = 'my_photo.jpg'  # Replace with your actual local image file name
SCOPES = [
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/drive' 
]

def append_image_to_doc():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    
    docs_service = build('docs', 'v1', credentials=creds)
    drive_service = build('drive', 'v3', credentials=creds)

    print(f"1. Uploading '{IMAGE_FILENAME}' to Drive...")

    # --- STEP 1: UPLOAD LOCAL IMAGE TO DRIVE ---
    file_metadata = {'name': IMAGE_FILENAME}
    media = MediaFileUpload(IMAGE_FILENAME, mimetype='image/jpeg')
    
    file = drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id, webContentLink'
    ).execute()
    
    file_id = file.get('id')
    # webContentLink is the direct download link Docs needs
    image_url = file.get('webContentLink') 

    print(f"   Uploaded! ID: {file_id}")

    # --- STEP 2: MAKE IMAGE READABLE ---
    # The Docs API needs permission to download this image to insert it.
    # We temporarily make it readable by "anyone" so the API server can fetch it.
    drive_service.permissions().create(
        fileId=file_id,
        body={'type': 'anyone', 'role': 'reader'},
        fields='id'
    ).execute()

    # --- STEP 3: FIND END OF DOCUMENT ---
    doc = docs_service.documents().get(documentId=EXISTING_DOC_ID).execute()
    content = doc.get('body').get('content')
    end_index = content[-1]['endIndex'] - 1

    print(f"2. Appending image at index {end_index}...")

    # --- STEP 4: INSERT IMAGE ---
    requests = [
        {
            'insertInlineImage': {
                'uri': image_url,
                'location': {
                    'index': end_index
                },
                'objectSize': {
                    'height': {'magnitude': 300, 'unit': 'PT'}, # Optional: Resize
                    'width': {'magnitude': 300, 'unit': 'PT'}
                }
            }
        }
    ]

    docs_service.documents().batchUpdate(
        documentId=EXISTING_DOC_ID, body={'requests': requests}).execute()

    print("Success! Image appended to the end of the document.")

    # Optional: You could delete the file from Drive here if you don't need to keep it,
    # as Docs makes its own internal copy once inserted.

if __name__ == '__main__':
    # Ensure you have a file named 'my_photo.jpg' (or whatever you set above)
    # in this folder before running.
    try:
        append_image_to_doc()
    except Exception as e:
        print(f"Error: {e}")