import io
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# --- CONFIGURATION ---
SERVICE_ACCOUNT_FILE = 'service_account.json'
EXISTING_DOC_ID = '10YTudX4LmIjxsoT2gOJFw2-yKMRYCxSJTNkHf4hZqoE'
# We need scopes to both READ (export) and WRITE (update)
SCOPES = ['https://www.googleapis.com/auth/drive']

def append_html_table_to_doc():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    
    drive_service = build('drive', 'v3', credentials=creds)

    print(f"Reading existing content from Doc ID: {EXISTING_DOC_ID}...")

    # 1. DOWNLOAD EXISTING CONTENT AS HTML
    # We export the current Google Doc to HTML format so we can append to it
    export_request = drive_service.files().export_media(
        fileId=EXISTING_DOC_ID,
        mimeType='text/html'
    )
    current_content = export_request.execute().decode('utf-8')

    # 2. PREPARE YOUR NEW TABLE HTML
    # Note: We added a <br> before the table for spacing
    new_table_html = """
    <br><br>
    <table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 500px;">
      <tr><th colspan="2" style="text-align:center;">About Me</th></tr>
      <tr><td><strong>Name</strong></td><td>Md. Sishir Rahman Siam (Nothing)</td></tr>
      <tr><td><strong>Role</strong></td><td>Programmer & Web Developer</td></tr>
      <tr><td><strong>Skills</strong></td><td>Python, Django, JavaScript, React, PostgreSQL, DSA</td></tr>
      <tr><td><strong>Profiles</strong></td><td>
          <a href="https://leetcode.com/u/shishirRsiam/">LeetCode</a> |
          <a href="https://codeforces.com/profile/shishirRsiam">Codeforces</a> |
          <a href="https://github.com/shishirRsiam">GitHub</a>
      </td></tr>
      <tr><td><strong>Timezone</strong></td><td>UTC+6</td></tr>
      <tr><td><strong>Email</strong></td><td>you@example.com</td></tr>
    </table>
    """

    # 3. MERGE CONTENTS
    # We look for the closing </body> tag and insert our table before it
    if "</body>" in current_content:
        updated_html = current_content.replace("</body>", f"{new_table_html}</body>")
    else:
        # Fallback if </body> is missing (rare)
        updated_html = current_content + new_table_html

    print("Uploading merged content...")

    # 4. OVERWRITE WITH THE MERGED HTML
    media = MediaIoBaseUpload(
        io.BytesIO(updated_html.encode('utf-8')),
        mimetype='text/html',
        resumable=True
    )

    drive_service.files().update(
        fileId=EXISTING_DOC_ID,
        media_body=media,
        # mimeType is implicitly handled, but setting it ensures conversion
        fields='id'
    ).execute()

    print("Success! Table added to existing content.")
    print(f"Check URL: https://docs.google.com/document/d/{EXISTING_DOC_ID}/edit")

if __name__ == '__main__':
    append_html_table_to_doc()
    