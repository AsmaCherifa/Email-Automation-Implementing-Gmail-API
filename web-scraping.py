import os.path
import pandas as pd
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def main():
   
    print("Starting Gmail API script...")
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    # Call the Gmail API
    try:
        service = build('gmail', 'v1', credentials=creds)
        print("Fetching messages...")
        results = service.users().messages().list(userId='me', maxResults=50).execute()
        messages = results.get('messages', [])
    except Exception as e:
        print(f'An error occurred: {e}')
        return

    emails_data = []

    if not messages:
        print('No messages found.')
    else:
        print(f'Found {len(messages)} messages.')
        for message in messages:
            msg = service.users().messages().get(userId='me', id=message['id']).execute()
            payload = msg['payload']
            headers = payload['headers']
            subject = ''
            sender = ''
            for header in headers:
                if header['name'] == 'Subject':
                    subject = header['value']
                if header['name'] == 'From':
                    sender = header['value']
            date = msg['internalDate']
            emails_data.append({
                'Sender': sender,
                'Subject': subject,
                'Received Time': date
            })

    print("Saving data to Excel...")
    df = pd.DataFrame(emails_data)
    df.to_excel('emails_data.xlsx', index=False)
    print("Data saved successfully.")

if __name__ == '__main__':
    main()
