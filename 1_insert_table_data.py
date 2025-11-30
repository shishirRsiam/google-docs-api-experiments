from google.oauth2 import service_account
from googleapiclient.discovery import build

SERVICE_ACCOUNT_FILE = "service_account.json"
DOC_ID = "1RxnVPcF1eOUv_buylyl8iX491ysDO8B-DwUpisLXDro"
DOC_URL = f"https://docs.google.com/document/d/1RxnVPcF1eOUv_buylyl8iX491ysDO8B-DwUpisLXDro/edit"



def add_student_data(document_id, all_data):
    scopes = ["https://www.googleapis.com/auth/documents"]
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=scopes
    )
    docs = build("docs", "v1", credentials=creds)

    # --- Step 1: Define your Data ---
    headers = ["Name", "Subject", "Grade"]
    students = [
        ["Sishir Smith", "Math", "F"],
        ["Bob Siam", "History", "B+"],
        ["John Doe", "English", "A"],
    ]
    
    # --- Step 2: Insert the Empty Table ---
    # We insert a 3x3 table at index 1 (start of doc)
    table_request = [
        {
            "insertTable": {
                "rows": len(all_data),
                "columns": len(all_data[0]),
                "location": {"index": 1}
            }
        }
    ]
    
    docs.documents().batchUpdate(
        documentId=document_id, 
        body={"requests": table_request}
    ).execute()

    print("Table inserted. Fetching document structure to find cell indices...")

    # --- Step 3: Fetch Document to find where the Cells are ---
    # We need to get the document again because the indices have changed
    doc = docs.documents().get(documentId=document_id).execute()
    content = doc.get('body').get('content')
    
    # Find the table we just created. 
    # Since we inserted at index 1, it is likely the second element (index 0 is usually section break)
    # We look for the first element that has a 'table' key.
    table = next(item for item in content if 'table' in item)
    
    # --- Step 4: Map Data to Cells ---
    # We flatten the table rows to make it easier to match with our data
    rows = table.get('table').get('tableRows')
    
    insert_text_requests = []
    
    # Prepare data list: Header + Student rows
    # all_data = [headers] + students
    

    # Loop through the rows and columns
    for r_idx, row in enumerate(rows):
        if r_idx < len(all_data):
            cells = row.get('tableCells')
            for c_idx, cell in enumerate(cells):
                if c_idx < len(all_data[r_idx]):
                    # The text we want to insert
                    text_to_add = all_data[r_idx][c_idx]
                    
                    # The specific index where this cell starts
                    # We add +1 to insert *inside* the cell, not before it
                    start_index = cell.get('startIndex')
                    
                    if start_index:
                        insert_text_requests.append({
                            "insertText": {
                                "text": text_to_add,
                                "location": {"index": start_index + 1}
                            }
                        })

    # --- Step 5: Execute Data Insertion (Reverse Order) ---
    # CRITICAL: We sort requests by index in DESCENDING order.
    # If we insert at index 10, index 20 becomes 21. 
    # If we insert backwards (20, then 10), the indices remain valid during the batch.
    # insert_text_requests.sort(key=lambda x: x['insertText']['location']['index'], reverse=True)

    insert_text_requests.reverse()

    if insert_text_requests:
        docs.documents().batchUpdate(
            documentId=document_id, 
            body={"requests": insert_text_requests}
        ).execute()
        print("Adding student data...")
        print(insert_text_requests)
        print("Student data added successfully!")
    else:
        print("No data to add.")

# Run the function


all_data = [
    ["নাম", "বয়স", "পিতা"],
    ["মোঃ ইসমাইল খান বাবু", "২৩", "আব্দুর রহমান"],
    ["মোঃ জাহিদ ওরফে তুহিন", "২৫", "মোঃ আঃ মজিদ"],
    ["মোঃ আলী হাসান", "৩০", "মোঃ মিজানুর রহমান"],
    ["মোঃ রুহুল আমিন", "২৮", "মোঃ সিদ্দিকুর রহমান"],
    ["মোঃ কামরুল ইসলাম", "২৭", "মোঃ আবুল কালাম"]
]

all_data = response_from_gpt()

add_student_data(DOC_ID, all_data)




[
    {"insertText": {"text": "A", "location": {"index": 30}}},
    {"insertText": {"text": "English", "location": {"index": 28}}},
    {"insertText": {"text": "John Doe", "location": {"index": 26}}},
    {"insertText": {"text": "B+", "location": {"index": 23}}},
    {"insertText": {"text": "History", "location": {"index": 21}}},
    {"insertText": {"text": "Bob Siam", "location": {"index": 19}}},
    {"insertText": {"text": "F", "location": {"index": 16}}},
    {"insertText": {"text": "Math", "location": {"index": 14}}},
    {"insertText": {"text": "Sishir Smith", "location": {"index": 12}}},
    {"insertText": {"text": "Grade", "location": {"index": 9}}},
    {"insertText": {"text": "Subject", "location": {"index": 7}}},
    {"insertText": {"text": "Name", "location": {"index": 5}}}
]

[
  {"insertText": {"text": "Name", "location": {"index": 5}}},
  {"insertText": {"text": "Subject", "location": {"index": 7}}},
  {"insertText": {"text": "Grade", "location": {"index": 9}}},
  {"insertText": {"text": "Sishir Smith", "location": {"index": 12}}},
  {"insertText": {"text": "Math", "location": {"index": 14}}},
  {"insertText": {"text": "F", "location": {"index": 16}}},
  {"insertText": {"text": "Bob Siam", "location": {"index": 19}}},
  {"insertText": {"text": "History", "location": {"index": 21}}},
  {"insertText": {"text": "B+", "location": {"index": 23}}},
  {"insertText": {"text": "John Doe", "location": {"index": 26}}},
  {"insertText": {"text": "English", "location": {"index": 28}}},
  {"insertText": {"text": "A", "location": {"index": 30}}}
]


