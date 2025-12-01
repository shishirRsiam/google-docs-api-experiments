import io
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# --- CONFIGURATION ---
SERVICE_ACCOUNT_FILE = 'service_account.json'
EXISTING_DOC_ID = '10YTudX4LmIjxsoT2gOJFw2-yKMRYCxSJTNkHf4hZqoE'
SCOPES = ['https://www.googleapis.com/auth/drive']

# The specific placeholder text to look for
PLACEHOLDER_TEXT = '<<replace_with_new_html_table>>'

def append_html_table_to_doc():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    
    print(f"Reading existing content from Doc ID: {EXISTING_DOC_ID}...")

    drive_service = build('drive', 'v3', credentials=creds)

    # 1. DOWNLOAD EXISTING CONTENT AS HTML
    export_request = drive_service.files().export_media(
        fileId=EXISTING_DOC_ID,
        mimeType='text/html'
    )
    current_content = export_request.execute().decode('utf-8')

    # Save backup of original (optional, for debugging)
    with io.open('backup_original.html', 'w', encoding='utf-8') as f:
        f.write(current_content)

    # 2. PREPARE YOUR NEW TABLE HTML
    new_table_html = """
    <table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 500px;">
      <tr><th colspan="2" style="text-align:center;">About Me</th></tr>
      <tr><td><strong>Name</strong></td><td>মোঃ শিশির রহমান সিয়াম</td></tr>
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

    # 3. REPLACE THE PLACEHOLDER
    # Google Docs exports symbols as HTML entities. 
    # << becomes &lt;&lt; and >> becomes &gt;&gt;
    # We define the escaped version to catch this.
    placeholder_escaped = PLACEHOLDER_TEXT.replace("<", "&lt;").replace(">", "&gt;")

    if placeholder_escaped in current_content:
        print(f"Found HTML-escaped placeholder '{placeholder_escaped}'. Replacing...")
        updated_html = current_content.replace(placeholder_escaped, new_table_html)
    elif PLACEHOLDER_TEXT in current_content:
        print(f"Found raw placeholder '{PLACEHOLDER_TEXT}'. Replacing...")
        updated_html = current_content.replace(PLACEHOLDER_TEXT, new_table_html)
    else:
        print("WARNING: Placeholder not found! Check 'backup_original.html' to see how the text looks.")
        # Fallback: Append to end if not found
        if "</body>" in current_content:
            updated_html = current_content.replace("</body>", f"<br><b>(Placeholder not found, appended instead):</b><br>{new_table_html}</body>")
        else:
            updated_html = current_content + new_table_html

    print("Uploading merged content...")

    # Save local copy of result (optional)
    with io.open('updated_content.html', 'w', encoding='utf-8') as f:
        f.write(updated_html)

    # 4. OVERWRITE WITH THE MERGED HTML
    media = MediaIoBaseUpload(
        io.BytesIO(updated_html.encode('utf-8')),
        mimetype='text/html',
        resumable=True
    )

    drive_service.files().update(
        fileId=EXISTING_DOC_ID,
        media_body=media,
        fields='id'
    ).execute()

    print("Success! Table swapped into document.")
    print(f"Check URL: https://docs.google.com/document/d/{EXISTING_DOC_ID}/edit")

if __name__ == '__main__':
    append_html_table_to_doc()

