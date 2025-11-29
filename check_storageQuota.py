from google.oauth2 import service_account
from googleapiclient.discovery import build

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

creds = service_account.Credentials.from_service_account_file(
    "service_account.json",
    scopes=SCOPES
)

drive_service = build("drive", "v3", credentials=creds)

about = drive_service.about().get(fields="storageQuota").execute()

print(about)

limit = int(about["storageQuota"]["limit"])
usage = int(about["storageQuota"]["usage"])
free = limit - usage

print("Total (MB):", limit / (1024 * 1024))
print("Used  (MB):", usage / (1024 * 1024))
print("Free  (MB):", free / (1024 * 1024))

