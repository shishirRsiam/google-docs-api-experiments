import io 
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

class GoogleDocs():
    _name = 'google.docs'
    _description = 'Custom Docs Helper Methods'

    def test_method(self, name="Shishir"):
        print('\n'*5)
        print(f'-> Hello {name}')
        creds = self.get_credentials()
        print('-> Credentials Received', creds)
        print('-> google Docs Initialized')

    def _build_docs_service(self, credentials=None):
        if credentials is None:
            credentials = self.get_credentials()
        return build('docs', 'v1', credentials=credentials)
    
    def _build_drive_service(self, credentials=None):
        if credentials is None:
            credentials = self.get_credentials()
        return build('drive', 'v3', credentials=credentials)

    def get_credentials(self):
        print('-> get_credentials called')
        return self.env['google.helper'].get_google_service_credentials()

    def get_html_content(self, docs_id, drive_service=None):
        if drive_service is None:
            drive_service = self._build_drive_service()
        export_request = drive_service.files().export_media(
            fileId=docs_id, mimeType='text/html'
        )
        return export_request.execute().decode('utf-8')
    
    def execute_batchUpdate(self, docs_id, requests, docs_service=None):
        if docs_service is None:
            docs_service = self._build_docs_service()
        try:
            docs_service.documents().batchUpdate(
                documentId=docs_id, body={"requests": requests}
            ).execute()

        except Exception as e:
            print('Error replacing text:', e)

    def merge_html(self, PLACEHOLDER_TEXT, current_html_content, new_html_content):
        placeholder_escaped = PLACEHOLDER_TEXT.replace("<", "&lt;").replace(">", "&gt;")

        if placeholder_escaped in current_html_content:
            print(f"Found HTML-escaped placeholder '{placeholder_escaped}'. Replacing...")
            merged_html = current_html_content.replace(placeholder_escaped, new_html_content)
        elif PLACEHOLDER_TEXT in current_html_content:
            print(f"Found raw placeholder '{PLACEHOLDER_TEXT}'. Replacing...")
            merged_html = current_html_content.replace(PLACEHOLDER_TEXT, new_html_content)
        else:
            print("WARNING: Placeholder not found! Appending to end instead...")
            # Fallback: Append to end if not found
            if "</body>" in current_html_content:
                merged_html = current_html_content.replace("</body>", f"<br><b>(Placeholder not found, appended to end instead):</b><br>{new_html_content}</body>")
            else:
                merged_html = current_html_content + new_html_content

        return merged_html

    def append_html(self, docs_id, replace_text="", new_html="", credentials=None):
        if not new_html:
            return False
        
        if credentials is None:
            credentials = self.get_credentials()

        if not docs_id:
            docs_id = '10YTudX4LmIjxsoT2gOJFw2-yKMRYCxSJTNkHf4hZqoE'

        drive_service = self._build_drive_service(credentials)
        current_content = self.get_html_content(docs_id, drive_service)
        merged_html = self.merge_html(replace_text, current_content, new_html)

        print("Uploading merged content...")

        self.update_html(docs_id, merged_html, drive_service)


    def update_html(self, docs_id, html_content, drive_service=None):
        print('-> update_html called')
        if drive_service is None:
            drive_service = self._build_drive_service()
        media = MediaIoBaseUpload(
            io.BytesIO(html_content.encode('utf-8')), mimetype='text/html', resumable=True
        )

        drive_service.files().update(
            fileId=docs_id, media_body=media, fields='id'
        ).execute()

        print("Success! Table added to existing content.")
        print(f"Check URL: https://docs.google.com/document/d/{docs_id}/edit")

        return docs_id

    def create_document(self, title, parent_id="1ST4IhB7lWNZqLQGZPS7FzhJembtB9rMY"):
        if parent_id:
            return self.create_in_folder(title, parent_id)
        request = {
            "title": title,
        }
        docs_service = self._build_docs_service()
        doc = docs_service.documents().create(body=request).execute()

        doc_id = doc["documentId"]
        url = "https://docs.google.com/document/d/%s/edit" % doc_id
        print(f'-> Document Created With ID: {doc_id}')
        print(f'-> Document URL: {url}')
        
        return doc_id, url

    def share_document(self, document_id, email=None, role="writer", credentials=None):
        if credentials is None:
            credentials = self.get_credentials()
        if email is None:
            email = "shishir.siam01@gmail.com" # DEV EMAIL
        permission = {
            "type": "user",
            "role": role,
            "emailAddress": email
        }
        drive_service = self._build_drive_service(credentials)
        drive_service.permissions().create(
            fileId=document_id, body=permission,
            sendNotificationEmail=False, fields="id"
        ).execute()
        print(f'-> Document Shared With: {email}')

    def copy_document(self, docs_id, file_name="New Document", parent_google_drive_id=None, credentials=None):
        if credentials is None:
            credentials = self.get_credentials()

        drive_service = build("drive", "v3", credentials=credentials)
        try:
            body = {"name": file_name}
            if parent_google_drive_id:
                body["parents"] = [parent_google_drive_id]

            copied_file = drive_service.files().copy(
                fileId=docs_id, body=body
            ).execute()

            copied_docs_id = copied_file['id']
            print('-> New Copied Document ID:', f'https://docs.google.com/document/d/{copied_docs_id}')
            return copied_docs_id

        except Exception as e:
            print('Error copying document:', e)
            return False
    
    def replace_text(self, docs_id, placeholder, replacement_text, credentials=None):
        if credentials is None:
            credentials = self.get_credentials()

        print("placeholder:", placeholder)         # will show 'Mahidul Islam Milon de'
        print('replacement_text', replacement_text)

        if isinstance(replacement_text, (tuple, list)) and len(replacement_text) >= 2:
            replacement_text = replacement_text[1]
        print('replacement_text', replacement_text)
        docs_service = build("docs", "v1", credentials=credentials)
        try:
            requests = [
                {
                    "replaceAllText": {
                        "containsText": {
                            "text": placeholder,
                            "matchCase": True
                        },
                        "replaceText": replacement_text
                    }
                }
            ]

            docs_service.documents().batchUpdate(
                documentId=docs_id, body={"requests": requests}
            ).execute()

        except Exception as e:
            print('Error replacing text:', e)

    def share_document_with_email(self, docs_id, credentials=None): # DEV TESTING
        if credentials is None:
            credentials = self.get_credentials()
        drive_service = build("drive", "v3", credentials=credentials)
        try:
            # email_address = [
            #     "shishir.siam01@gmail.com",
            # ]  
            
            email_address = [
                "shishir.siam01@gmail.com"
            ]  

            for email in email_address:
                drive_service.permissions().create(
                    fileId=docs_id,
                    body={"type": "user", "role": "writer", "emailAddress": email}
                ).execute()
                print()
                print('-> Document shared with:', email)
        
        except Exception as e:
            print('Error sharing document:', e)

    def get_google_docs_content(self, docs_id, credentials=None):
        service = build("docs", "v1", credentials=credentials)
        try:
            doc = service.documents().get(documentId=docs_id).execute()
            content = self.extract_text_from_google_docs(doc)
            return content
        except Exception as e:
            print('Error:', e)
            return False

    def extract_text_from_google_docs(self, doc):
        text = []
        for element in doc.get("body", {}).get("content", []):
            if "paragraph" in element:
                for run in element["paragraph"].get("elements", []):
                    if "textRun" in run:
                        text.append(run["textRun"]["content"])
        return "".join(text)

    def get_document(self, document_id):
        docs_service = self._build_docs_service()
        return docs_service.documents().get(documentId=document_id).execute()

    def insert_text(self, document_id, text, index=1):
        request = {
            "insertText": {
                "location": {"index": index},
                "text": text
            }
        }
        docs_service = self._build_docs_service()
        docs_service.documents().batchUpdate(
            documentId=document_id,
            body={"requests": [request]}
        ).execute()

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

        docs_service = self._build_docs_service()
        docs_service.documents().batchUpdate(
            documentId=document_id,
            body={"requests": requests}
        ).execute()

    def create_in_folder(self, title, folder_id):
        metadata = {
            "name": title,
            "mimeType": "application/vnd.google-apps.document",
            "parents": [folder_id]
        }
        drive_service = self._build_drive_service()
        file = drive_service.files().create(body=metadata, fields="id").execute()
        doc_id = file["id"]
        url = "https://docs.google.com/document/d/%s/edit" % doc_id
        return doc_id, url

    def update_title(self, document_id, new_title):
        drive_service = self._build_drive_service()
        drive_service.files().update(
            fileId=document_id,
            body={"name": new_title}
        ).execute()

    def delete_document(self, document_id):
        drive_service = self._build_drive_service()
        drive_service.files().delete(fileId=document_id).execute()

    def bold_text(self, document_id, target):
        docs_service = self._build_docs_service()
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
            docs_service.documents().batchUpdate(
                documentId=document_id,
                body={"requests": requests}
            ).execute()

    def delete_text(self, document_id, start_index, end_index):
        docs_service = self._build_docs_service()
        request = {
            "deleteContentRange": {
                "range": {
                    "startIndex": start_index,
                    "endIndex": end_index
                }
            }
        }
        docs_service.documents().batchUpdate(
            documentId=document_id,
            body={"requests": [request]}
        ).execute()

    def insert_page_break(self, document_id):
        docs_service = self._build_docs_service()
        doc = self.get_document(document_id)
        end_index = doc["body"]["content"][-1]["endIndex"] - 1

        request = {
            "insertPageBreak": {
                "location": {"index": end_index}
            }
        }
        docs_service.documents().batchUpdate(
            documentId=document_id,
            body={"requests": [request]}
        ).execute()

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

        docs_service.documents().batchUpdate(
            documentId=document_id,
            body={"requests": requests}
        ).execute()

        # Apply bullet formatting
        end = start + len(text)

        docs_service = self._build_docs_service()
        docs_service.documents().batchUpdate(
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


    # def save_data(self, data, file_name="output.txt"):
    #     import time 
    #     cur_time = time.strftime("%d-%m-%Y-%H-%M-%S", time.localtime())
    #     append_text = f"[{cur_time}]\n{data}\n\n\n{'**'*50}\n\n\n\n"
    #     with open(file_name, 'a') as f:
    #         f.write(append_text)
