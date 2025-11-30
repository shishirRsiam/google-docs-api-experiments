
html_text = "<b><i>Md. Sishir Rahman Siam</i></b>"

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import io

SERVICE_ACCOUNT_FILE = "service_account.json"
# We need the DRIVE scope to create/upload files
SCOPES = ['https://www.googleapis.com/auth/drive']

def create_doc_from_html():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    
    # Use the DRIVE service, not Docs
    drive_service = build('drive', 'v3', credentials=creds)

    # Your HTML content
    html_content = """<table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 500px;">
  <tr>
    <th colspan="2" style="text-align:center;">About Me</th>
  </tr>

  <tr>
    <td><strong>Name</strong></td>
    <td>Md. Sishir Rahman Siam (shishirRsiam)</td>
  </tr>

  <tr>
    <td><strong>Role</strong></td>
    <td>Programmer & Web Developer</td>
  </tr>

  <tr>
    <td><strong>Skills</strong></td>
    <td>Python, Django, JavaScript, React, PostgreSQL, DSA</td>
  </tr>

  <tr>
    <td><strong>Profiles</strong></td>
    <td>
      <a href="https://leetcode.com/u/shishirRsiam/">LeetCode</a> |
      <a href="https://codeforces.com/profile/shishirRsiam">Codeforces</a> |
      <a href="https://github.com/shishirRsiam">GitHub</a>
    </td>
  </tr>

  <tr>
    <td><strong>Timezone</strong></td>
    <td>UTC+6</td>
  </tr>

  <tr>
    <td><strong>Email</strong></td>
    <td>you@example.com</td>
  </tr>
</table>
"""
    
    file_metadata = {
        'name': 'My Converted HTML Doc',
        # This mimeType tells Google to convert the uploaded file into a Google Doc
        'mimeType': 'application/vnd.google-apps.document'
    }

    # Prepare the HTML as a file stream
    media = MediaIoBaseUpload(
        io.BytesIO(html_content.encode('utf-8')),
        mimetype='text/html',
        resumable=True
    )

    # Create the file
    file = drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id'
    ).execute()

    drive_service.permissions().create(
            fileId=file.get('id'),
            body={
                "type": "user",
                "role": "writer",
                "emailAddress": 'shishir.siam01@gmail.com',
            },
            sendNotificationEmail=False,
            fields="id"
        ).execute()

    print(f"Created new Doc ID: {file.get('id')}")
    print(f"File URL: https://docs.google.com/document/d/{file.get('id')}/edit")

if __name__ == '__main__':
    create_doc_from_html()