from google.oauth2 import service_account
from googleapiclient.discovery import build

# --- Configuration ---
SHARE_EMAIL = "shishir.siam01@gmail.com"
DOC_TITLE = "API Service Account Document"
INSERT_TEXT = "Hello Sishir (Created via Service Account)"

SCOPES = [
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/drive"
]

SERVICE_ACCOUNT_FILE = "service_account.json"


def main():
    print("🚀 Loading service account credentials...")
    try:
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE,
            scopes=SCOPES
        )
        print("✅ Credentials loaded.")
    except Exception as e:
        print(f"❌ Failed to load credentials: {e}")
        return

    # Build service clients
    docs_service = build("docs", "v1", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)

    # 1. Create Doc
    print("\n📝 Creating Google Doc with service account...")
    try:
        doc = docs_service.documents().create(
            body={"title": DOC_TITLE}
        ).execute()

        document_id = doc.get("documentId")
        doc_url = f"https://docs.google.com/document/d/{document_id}/edit"

        print(f"✅ Document Created!")
        print(f"📄 ID: {document_id}")
        print(f"🔗 URL: {doc_url}")

    except Exception as e:
        print(f"❌ Error creating doc: {e}")
        return

    # 2. Share Doc
    print(f"\n🤝 Sharing with: {SHARE_EMAIL}")
    try:
        drive_service.permissions().create(
            fileId=document_id,
            body={
                "type": "user",
                "role": "writer",
                "emailAddress": SHARE_EMAIL,
            },
            sendNotificationEmail=False,
            fields="id"
        ).execute()

        print("✅ Shared successfully.")

    except Exception as e:
        print(f"❌ Sharing failed: {e}")

    # 3. Insert Text
    print(f"\n✍️ Inserting text: {INSERT_TEXT}")
    try:
        requests = [
            {
                "insertText": {
                    "location": {"index": 1},
                    "text": INSERT_TEXT
                }
            }
        ]

        docs_service.documents().batchUpdate(
            documentId=document_id,
            body={"requests": requests}
        ).execute()

        print("✅ Text inserted.")

    except Exception as e:
        print(f"❌ Error inserting text: {e}")

    print("\n✨ Complete.")


if __name__ == "__main__":
    main()
