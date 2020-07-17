import time
import datetime
import base64
import email
import pickle
import os.path
from apiclient import errors
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://mail.google.com/']

CALENDAR_SCOPES = ['https://www.googleapis.com/auth/calendar']


def get_gmail_service():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return build('gmail', 'v1', credentials=creds)


def get_latest_message_id(service, user):
    try:
        # Get latest message
        response = service.users().messages().list(userId=user, maxResults=1,
                                                   q="from:(receipt@bibliotheca.com) subject:Kvittering").execute()
        # Get message id
        if 'messages' in response:
            return response['messages'][0]['id']
        else:
            return None

    except errors.HttpError:
        print("An error occured while fetching the message id")


def get_books_return_date(service, user, msgId):
    try:
        message = service.users().messages().get(
            userId=user, id=msgId, format="full").execute()

        msg_str = base64.urlsafe_b64decode(
            message['payload']['body']['data'].encode('utf-8'))

        mime_msg = email.message_from_string(msg_str.decode("utf-8"))

        # Split mime_msg in order to extract date for return of books
        split_message = mime_msg.as_string().splitlines()

        # Loop over message and return date of return
        returnDate = None
        for text in split_message:
            if text.startswith("Afleveres"):
                returnDate = text.split(": ")[1].strip()

        return returnDate

    except errors.HttpError:
        print("An error occured while trying to fetch message")


def get_calendar_service():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('calendar-token.pickle'):
        with open('calendar-token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', CALENDAR_SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('calendar-token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return build('calendar', 'v3', credentials=creds)


def create_cal_event(service, furtureDate, nextDay):
    # Specify event
    event = {
        'summary': 'Afl. biblioteksbøger',
        'start': {
            'dateTime': furtureDate,
            'timeZone': 'Europe/Copenhagen',
        },
        'end': {
            'dateTime': nextDay,
            'timeZone': 'Europe/Copenhagen',
        },
    }

    return service.events().insert(calendarId='primary', body=event).execute()


def trash_message(service, user, msgId):
    try:
        service.users().messages().trash(userId=user, id=msgId).execute()
    except errors.HttpError:
        print("An error occured while trashing message")


gmail_service = get_gmail_service()
message_id = get_latest_message_id(gmail_service, 'me')

if message_id != None:
    # Get book return date from message
    book_return_date = get_books_return_date(gmail_service, 'me', message_id)

    # Get calendar service
    calender_service = get_calendar_service()

    # Create future dates
    return_date_obj = datetime.datetime(
        int(book_return_date[6:]), int(book_return_date[3:5]), int(book_return_date[:2]))
    future_return_date = return_date_obj + datetime.timedelta(days=30)
    nextDay = future_return_date + datetime.timedelta(days=1)

    # Create calendar event
    event = create_cal_event(calender_service, str(
        future_return_date.isoformat()), str(nextDay.isoformat()))

    # Delete message if an event has been created
    if event != None:
        trash_message(gmail_service, 'me', message_id)
