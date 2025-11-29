from google.oauth2 import service_account
from googleapiclient.discovery import build

SERVICE_ACCOUNT_FILE = "service_account.json"
DOC_ID = "1RxnVPcF1eOUv_buylyl8iX491ysDO8B-DwUpisLXDro"  # replace with your existing doc ID
DOC_URL = f"https://docs.google.com/document/d/{DOC_ID}/edit"

def replace_text_in_doc(document_id, old_text, new_text):
    # Authenticate
    scopes = ["https://www.googleapis.com/auth/documents"]
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=scopes
    )

    docs_service = build("docs", "v1", credentials=creds)

    # Replace text
    requests = [
        {
            "replaceAllText": {
                "containsText": {
                    "text": old_text,
                    "matchCase": True
                },
                "replaceText": new_text
            }
        }
    ]

    docs_service.documents().batchUpdate(
        documentId=document_id,
        body={"requests": requests}
    ).execute()

    print("✅ Replaced '%s' with '%s' in document %s" % (old_text, new_text, document_id))


if __name__ == "__main__":
    replace_text_in_doc(DOC_ID, "opt/bu", "new_text_here")


from google.oauth2 import service_account
from googleapiclient.discovery import build

SERVICE_ACCOUNT_FILE = "service_account.json"
DOC_ID = "1RxnVPcF1eOUv_buylyl8iX491ysDO8B-DwUpisLXDro"  # replace with doc ID
DOC_URL = f"https://docs.google.com/document/d/{DOC_ID}/edit"


# -----------------------------
# FUNCTION: INSERT TABLE
# -----------------------------
def insert_table_in_doc(document_id):
    scopes = ["https://www.googleapis.com/auth/documents"]
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=scopes
    )

    docs_service = build("docs", "v1", credentials=creds)

    table_requests = {
        "requests": [
            {
                "insertTable": {
                    "rows": 10,
                    "columns": 2,
                    "location": {
                        "index": 1
                    }
                }
            },

            { "insertText": { "location": { "index": 2 },  "text": "ক্রমিক" }},
            { "insertText": { "location": { "index": 3 },  "text": "বিবরণ" }},

            { "insertText": { "location": { "index": 6 },  "text": "১ম জব্দতারিখ" }},
            { "insertText": { "location": { "index": 7 },  "text": "০৮ জুলাই, ২০২৫" }},

            { "insertText": { "location": { "index": 10 }, "text": "সময়" }},
            { "insertText": { "location": { "index": 11 }, "text": "১২:৫০ ঘটিকা" }},

            { "insertText": { "location": { "index": 14 }, "text": "প্রস্তুতকারী কর্মকর্তা" }},
            { "insertText": { "location": { "index": 15 }, "text": "সাব-ইন্সপেক্টর (নিঃ) নয়ন কুমার চক্রবর্তী" }},

            { "insertText": { "location": { "index": 18 }, "text": "বিপি নম্বর" }},
            { "insertText": { "location": { "index": 19 }, "text": "BP-7999015253" }},

            { "insertText": { "location": { "index": 22 }, "text": "যন্ত্রের ধরণ" }},
            { "insertText": { "location": { "index": 23 }, "text": "DVR মেশিন" }},

            { "insertText": { "location": { "index": 26 }, "text": "ব্র্যান্ড/কোম্পানি" }},
            { "insertText": { "location": { "index": 27 }, "text": "DHUA" }},

            { "insertText": { "location": { "index": 30 }, "text": "হার্ড ডিস্ক ক্ষমতা" }},
            { "insertText": { "location": { "index": 31 }, "text": "4TB" }},

            { "insertText": { "location": { "index": 34 }, "text": "হার্ড ডিস্ক সিরিয়াল নম্বর (S/N)" }},
            { "insertText": { "location": { "index": 35 }, "text": "WCC7K3FAVX27" }},

            { "insertText": { "location": { "index": 38 }, "text": "হার্ড ডিস্ক মডেল (MDL)" }},
            { "insertText": { "location": { "index": 39 }, "text": "WD40PURX-69N69Y0" }},

            { "insertText": { "location": { "index": 42 }, "text": "ভিডিওর বিবরণ" }},
            { "insertText": { "location": { "index": 43 }, "text": "ঘটনাস্থলে যাতায়াতের ভিডিও সংরক্ষিত আছে" }}
        ]
    }

    # Execute API call
    docs_service.documents().batchUpdate(
        documentId=document_id,
        body=table_requests
    ).execute()

    print("✅ Table inserted into the document!")


# -----------------------------
# FUNCTION: REPLACE TEXT
# -----------------------------
def replace_text_in_doc(document_id, old_text, new_text):
    scopes = ["https://www.googleapis.com/auth/documents"]
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=scopes
    )
    docs_service = build("docs", "v1", credentials=creds)

    requests = [
        {
            "replaceAllText": {
                "containsText": {
                    "text": old_text,
                    "matchCase": True
                },
                "replaceText": new_text
            }
        }
    ]

    docs_service.documents().batchUpdate(
        documentId=document_id,
        body={"requests": requests}
    ).execute()

    print(f"🔄 Replaced '{old_text}' with '{new_text}'")


# -----------------------------
# MAIN EXECUTION
# -----------------------------
if __name__ == "__main__":
    insert_table_in_doc(DOC_ID)   # Insert table
    # replace_text_in_doc(DOC_ID, "opt/bu", "new_text_here")   # Optional text replace


