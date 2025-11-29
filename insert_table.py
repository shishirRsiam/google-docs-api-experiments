from google.oauth2 import service_account
from googleapiclient.discovery import build

SERVICE_ACCOUNT_FILE = "service_account.json"
DOC_ID = "1RxnVPcF1eOUv_buylyl8iX491ysDO8B-DwUpisLXDro"


def insert_table_in_doc(document_id):
    scopes = ["https://www.googleapis.com/auth/documents"]
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=scopes
    )

    docs_service = build("docs", "v1", credentials=creds)

    requests = {
        "requests": [
            {
                "insertTable": {
                    "rows": 3,
                    "columns": 3,
                    "location": {"index": 1}
                }
            }
        ]
    }

    print("📝 Inserting table into the document...")

    docs_service.documents().batchUpdate(
        documentId=document_id,
        body=requests
    ).execute()

    print("✅ Table inserted!")


if __name__ == "__main__":
    insert_table_in_doc(DOC_ID)
