import requests
import sys
import os

def get_conversation(api_token, convo_id):
    url = f'https://api2.frontapp.com/conversations/{convo_id}'
    headers = {
        'Accept': 'application/json',
        'Authorization': f'Bearer {api_token}'
    }
    print(f"Fetching conversation ID: {convo_id}")
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        conversation = response.json()
        print(f"Found conversation: {conversation.get('subject', 'No subject')}")
        return conversation
    else:
        print(f"Failed to retrieve conversation. Status code: {response.status_code}")
        print(response.text)
        return None

def get_messages(api_token, convo_id):
    url = f'https://api2.frontapp.com/conversations/{convo_id}/messages'
    headers = {
        'Accept': 'application/json',
        'Authorization': f'Bearer {api_token}'
    }
    print(f"Fetching messages for conversation ID: {convo_id}")
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        messages = response.json().get('_results', [])
        print(f"Found {len(messages)} messages in conversation {convo_id}")
        return messages
    else:
        print(f"Failed to retrieve messages for conversation {convo_id}. Status code: {response.status_code}")
        print(response.text)
        return []

def find_excel_attachments(messages):
    excel_files = []
    for message in messages:
        for attachment in message.get('attachments', []):
            if attachment['filename'].endswith(('.xls', '.xlsx')):
                download_link = attachment['url'] if 'url' in attachment else None
                excel_files.append({
                    'filename': attachment['filename'],
                    'url': download_link,
                    'id': attachment['id']
                })
    return excel_files

def sanitize_filename(filename):
    sanitized = "".join([c for c in filename if c.isalpha() or c.isdigit() or c in (' ', '.', '_')])
    return sanitized.rstrip()

def download_attachment(api_token, attachment, save_dir):
    url = attachment['url']
    headers = {
        'Authorization': f'Bearer {api_token}'
    }
    print(f"Downloading attachment {attachment['filename']} from {url}")
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        sanitized_filename = sanitize_filename(attachment['filename'])
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)
        filepath = os.path.join(save_dir, sanitized_filename)
        with open(filepath, 'wb') as f:
            f.write(response.content)
        print(f"Downloaded {sanitized_filename} to {filepath}")
    else:
        print(f"Failed to download attachment {attachment['filename']}. Status code: {response.status_code}")
        print(response.text)

def main():
    api_token = 'eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzY29wZXMiOlsic2NpbSIsInByb3Zpc2lvbmluZyIsInByaXZhdGU6KiIsInNoYXJlZDoqIiwia2IiLCJ0aW06NTYzOTA3MyJdLCJpYXQiOjE3MTc0Mzk5ODYsImlzcyI6ImZyb250Iiwic3ViIjoiMWJkNjBiZTFjMTc1MDVlNjQ0Y2QiLCJqdGkiOiI4MGZmYjk1NDkyM2I2ZDdlIn0.8q5l53xCueU3OU5I0_amFfpk_tgVZ1zNeYeBCeSJqnY'
    convo_id = '101025750618'
    save_dir = 'Reports'
    sys.stdout.reconfigure(encoding='utf-8')
    conversation = get_conversation(api_token, convo_id)
    if not conversation:
        print(f"Conversation with ID {convo_id} not found.")
        return
    messages = get_messages(api_token, convo_id)
    if not messages:
        print(f"No messages found in conversation {convo_id}")
        return
    excel_files = find_excel_attachments(messages)
    if excel_files:
        print("Excel files found:")
        for file in excel_files:
            print(f"Filename: {file['filename']}, Download URL: {file['url']}")
            download_attachment(api_token, file, save_dir)
    else:
        print("No Excel files found.")

if __name__ == "__main__":
    main()
