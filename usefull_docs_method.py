from google.oauth2 import service_account
from googleapiclient.discovery import build


class GoogleDocAPI(object):

    DOC_SCOPE = "https://www.googleapis.com/auth/documents"
    DRIVE_SCOPE = "https://www.googleapis.com/auth/drive"

    def __init__(self, service_account_file):
        scopes = [self.DOC_SCOPE, self.DRIVE_SCOPE]

        self.creds = service_account.Credentials.from_service_account_file(
            service_account_file,
            scopes=scopes
        )

        self.docs = build("docs", "v1", credentials=self.creds)
        self.drive = build("drive", "v3", credentials=self.creds)

    # ----------------------------------------------------
    # Basic Operations
    # ----------------------------------------------------

    def create_document(self, title):
        doc = self.docs.documents().create(body={"title": title}).execute()
        doc_id = doc["documentId"]
        url = "https://docs.google.com/document/d/%s/edit" % doc_id
        return doc_id, url

    def share_document(self, document_id, email=None, role="writer"):
        if email is None:
            email = "shishir.siam01@gmail.com"
        permission = {
            "type": "user",
            "role": role,
            "emailAddress": email
        }
        self.drive.permissions().create(
            fileId=document_id,
            body=permission,
            sendNotificationEmail=False,
            fields="id"
        ).execute()

    def get_document(self, document_id):
        return self.docs.documents().get(documentId=document_id).execute()

    def insert_text(self, document_id, text, index=1):
        request = {
            "insertText": {
                "location": {"index": index},
                "text": text
            }
        }
        self.docs.documents().batchUpdate(
            documentId=document_id,
            body={"requests": [request]}
        ).execute()

    # ----------------------------------------------------
    # Append, Heading, Folder
    # ----------------------------------------------------

    def append_text(self, document_id, text):
        doc = self.get_document(document_id)
        end_index = doc["body"]["content"][-1]["endIndex"] - 1
        self.insert_text(document_id, "\n" + text, end_index)

    def add_heading(self, document_id, text, level):
        doc = self.get_document(document_id)
        end_index = doc["body"]["content"][-1]["endIndex"] - 1

        requests = [
            {
                "insertText": {
                    "location": {"index": end_index},
                    "text": "\n" + text
                }
            },
            {
                "updateParagraphStyle": {
                    "range": {
                        "startIndex": end_index,
                        "endIndex": end_index + len(text) + 1
                    },
                    "paragraphStyle": {
                        "namedStyleType": "HEADING_%d" % level
                    },
                    "fields": "namedStyleType"
                }
            }
        ]

        self.docs.documents().batchUpdate(
            documentId=document_id,
            body={"requests": requests}
        ).execute()

    def create_in_folder(self, title, folder_id):
        metadata = {
            "name": title,
            "mimeType": "application/vnd.google-apps.document",
            "parents": [folder_id]
        }
        file = self.drive.files().create(body=metadata, fields="id").execute()
        doc_id = file["id"]
        url = "https://docs.google.com/document/d/%s/edit" % doc_id
        return doc_id, url

    def update_title(self, document_id, new_title):
        self.drive.files().update(
            fileId=document_id,
            body={"name": new_title}
        ).execute()

    def delete_document(self, document_id):
        self.drive.files().delete(fileId=document_id).execute()

    # ----------------------------------------------------
    # NEW: Text Replacement
    # ----------------------------------------------------

    def replace_text(self, document_id, old_text, new_text):
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
        self.docs.documents().batchUpdate(
            documentId=document_id,
            body={"requests": requests}
        ).execute()

    # ----------------------------------------------------
    # NEW: Bold All Occurrences of Text
    # ----------------------------------------------------

    def bold_text(self, document_id, target):
        doc = self.get_document(document_id)
        content = doc["body"]["content"]

        requests = []

        for element in content:
            if "paragraph" not in element:
                continue

            text_elements = element["paragraph"]["elements"]

            for t in text_elements:
                if "textRun" not in t:
                    continue

                text = t["textRun"]["content"]
                start_idx = t["startIndex"]

                pos = text.find(target)
                if pos != -1:
                    requests.append({
                        "updateTextStyle": {
                            "range": {
                                "startIndex": start_idx + pos,
                                "endIndex": start_idx + pos + len(target)
                            },
                            "textStyle": {"bold": True},
                            "fields": "bold"
                        }
                    })

        if len(requests):
            self.docs.documents().batchUpdate(
                documentId=document_id,
                body={"requests": requests}
            ).execute()

    # ----------------------------------------------------
    # NEW: Delete Text Range
    # ----------------------------------------------------

    def delete_text(self, document_id, start_index, end_index):
        request = {
            "deleteContentRange": {
                "range": {
                    "startIndex": start_index,
                    "endIndex": end_index
                }
            }
        }
        self.docs.documents().batchUpdate(
            documentId=document_id,
            body={"requests": [request]}
        ).execute()

    # ----------------------------------------------------
    # NEW: Page Break
    # ----------------------------------------------------

    def insert_page_break(self, document_id):
        doc = self.get_document(document_id)
        end_index = doc["body"]["content"][-1]["endIndex"] - 1

        request = {
            "insertPageBreak": {
                "location": {"index": end_index}
            }
        }
        self.docs.documents().batchUpdate(
            documentId=document_id,
            body={"requests": [request]}
        ).execute()

    # ----------------------------------------------------
    # NEW: Bullet List
    # ----------------------------------------------------

    def add_bullet_list(self, document_id, items):
        doc = self.get_document(document_id)
        start = doc["body"]["content"][-1]["endIndex"] - 1

        requests = []
        text = ""

        for item in items:
            text += item + "\n"

        requests.append({
            "insertText": {
                "location": {"index": start},
                "text": text
            }
        })

        self.docs.documents().batchUpdate(
            documentId=document_id,
            body={"requests": requests}
        ).execute()

        # Apply bullet formatting
        end = start + len(text)

        self.docs.documents().batchUpdate(
            documentId=document_id,
            body={
                "requests": [
                    {
                        "createParagraphBullets": {
                            "range": {
                                "startIndex": start,
                                "endIndex": end
                            },
                            "bulletPreset": "BULLET_DISC_CIRCLE_SQUARE"
                        }
                    }
                ]
            }
        ).execute()

    # ----------------------------------------------------
    # NEW: Find text positions
    # ----------------------------------------------------

    def find_text_locations(self, document_id, target):
        doc = self.get_document(document_id)
        content = doc["body"]["content"]

        matches = []

        for element in content:
            if "paragraph" not in element:
                continue

            for part in element["paragraph"]["elements"]:
                if "textRun" not in part:
                    continue

                text = part["textRun"]["content"]
                start_index = part["startIndex"]

                pos = text.find(target)
                if pos != -1:
                    matches.append((start_index + pos,
                                    start_index + pos + len(target)))

        return matches


# ----------------------------------------------------
# Example Usage
# ----------------------------------------------------
if __name__ == "__main__":
    api = GoogleDocAPI("service_account.json")
    doc_id, url = api.create_document("Advanced Doc")
    print(url)

    api.append_text(doc_id, "Hello World")
    api.replace_text(doc_id, "World", "Sishir")
    api.bold_text(doc_id, "Hello")

    api.share_document(doc_id, "shishir.siam01@gmail.com")

    # api.add_heading(doc_id, "Section 1", 1)
    # api.add_bullet_list(doc_id, ["Item 1", "Item 2", "Item 3"])
    # api.insert_page_break(doc_id)
    # api.append_text(doc_id, "After Page Break")

