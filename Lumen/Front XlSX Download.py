import os
import requests
import openpyxl
from openpyxl.utils.exceptions import InvalidFileException
from datetime import datetime

# Fetch environment variables
API_TOKEN = 'eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzY29wZXMiOlsic2NpbSIsInByb3Zpc2lvbmluZyIsInNoYXJlZDoqIiwicHJpdmF0ZToqIiwia2IiLCJ0aW06NTYzOTA3MyJdLCJpYXQiOjE3MjA2MzE2MDQsImlzcyI6ImZyb250Iiwic3ViIjoiMWJkNjBiZTFjMTc1MDVlNjQ0Y2QiLCJqdGkiOiJmYTE3ZmM3NGM3Y2EyZTBmIn0.0indI8nPShOUVncJojMcilUp4NDTUYy3HLmYABLNRPw'
INBOX_ID = 'inb_8wrvl'
DOWNLOAD_DIR = '/Users/cash/desktop/FrontExcel/Lumen/'
# Downloads/

if not API_TOKEN:
    raise ValueError("No API token provided. Please set the API token.")

headers = {
    'Authorization': f'Bearer {API_TOKEN}',
    'Accept': 'application/json'
}

session = requests.Session()
session.headers.update(headers)

def sanitize_string(s):
    invalid_chars = "<>:\"/\\|?*"
    for char in invalid_chars:
        s = s.replace(char, '_')
    return s.encode('ascii', 'ignore').decode('ascii')

def get_conversations(inbox_id, start_date, end_date):
    url = f'https://api2.frontapp.com/inboxes/{inbox_id}/conversations'
    params = {
        'q[start]': start_date.isoformat() + 'Z',
        'q[end]': end_date.isoformat() + 'Z'
    }
    response = session.get(url, params=params)
    response.raise_for_status()
    return response.json().get('_results', [])

def get_messages(conversation_id):
    url = f'https://api2.frontapp.com/conversations/{conversation_id}/messages'
    response = session.get(url)
    response.raise_for_status()
    return response.json().get('_results', [])

def download_attachment(attachment):
    url = attachment.get('url')
    filename = sanitize_string(attachment.get('filename', ''))
    if not filename:
        print("Skipping empty filename.")
        return
    
    if not (filename.endswith('.xls') or filename.endswith('.xlsx')):
        print(f"Skipping non-excel file: {filename}")
        return

    response = session.get(url)
    response.raise_for_status()
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    file_path = os.path.join(DOWNLOAD_DIR, sanitize_string(filename))

    with open(file_path, 'wb') as f:
        f.write(response.content)
    print(f'Downloaded {sanitize_string(filename)} to {file_path}')

def main(inbox_id):
    # Define the date range for filtering
    start_date = datetime(2024, 6, 10)
    end_date = datetime(2024, 6, 15)

    try:
        print(f"Fetching conversations for Inbox ID: {inbox_id} between {start_date} and {end_date}...")
        conversations = get_conversations(inbox_id, start_date, end_date)
        for conversation in conversations:
            conversation_id = conversation['id']
            print(f"Processing conversation ID: {conversation_id}...")
            messages = get_messages(conversation_id)
            found_attachments = False
            for message in messages:
                if message.get('subject') == "Sales by Product/Service Detail - FSM Copy":
                    if 'attachments' in message:
                        found_attachments = True
                        for attachment in message['attachments']:
                            print(f"Found attachment: {sanitize_string(attachment['filename'])}")
                            download_attachment(attachment)
            if not found_attachments:
                print("No attachments found in messages.")
    except requests.exceptions.HTTPError as err:
        print(f"HTTP error occurred: {err}")
    except Exception as err:
        print(f"An error occurred: {err}")

if __name__ == '__main__':
    main(INBOX_ID)
