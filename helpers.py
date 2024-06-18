SCOPES = [
    "https://www.googleapis.com/auth/calendar.readonly",
    "https://www.googleapis.com/auth/calendar.events",
    "https://www.googleapis.com/auth/gmail.compose",
    "https://www.googleapis.com/auth/gmail.readonly"]

from llama_index.readers.google import GmailReader, GoogleCalendarReader
from typing import Any, List, Optional, Union
from llama_index.core import  Document
from tabulate import tabulate
from IPython.display import display, Markdown
from bs4 import BeautifulSoup, MarkupResemblesLocatorWarning
import warnings
from datetime import datetime
from googleapiclient.discovery import build
import os

from google_auth_oauthlib.flow import InstalledAppFlow
import json

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

from email import message_from_string

warnings.filterwarnings("ignore", category=MarkupResemblesLocatorWarning)

def get_system_prompt():
  prompt = """
  You are Chris Nesbitt-Smith's executive assistant, your name is ai.cns.me
  You will answer all his queries very helpfully and concisely.
  Emojis are good if used appropriatetly.
  You have access to all his calendar events for the past three months and three months in the future to help you with this.
  You should format your responses in markdown, feel free to include quotes from relevant content to backup your response.
  You should address him in your response.
  You can provide hyperlinks to your sources in your response if formatted in markdown.
  Keep trying until you have a confident answer
  the following terms are too common, so you should avoid using them to search with:
    - catch up
    - meeting
    - call
  """
  return prompt


def ask_a_question(index, question: str):
  query_engine = index.as_query_engine(
    #  response_mode="compact", 
     similarity_top_k=20)
  # TODO: add current date and time to prompt
  # TODO: add current location to prompt
#   prompt = get_system_prompt()
#   response = query_engine.query(prompt + question)
  response = query_engine.query(question)
  format_answer(response)


def format_answer(response: any):
  display(Markdown(response.response))

  emails = []
  events = []
  for source_node in response.source_nodes:
    temp_node = source_node.copy()

    if 'url' in temp_node.metadata and 'Subject' in temp_node.metadata:
      temp_node.metadata['Subject'] = f"[{temp_node.metadata['Subject']}]({temp_node.metadata.pop('url')})"

    del temp_node.metadata['id']

    if 'DocType' in source_node.metadata and source_node.metadata['DocType'] == 'email':
      del temp_node.metadata['DocType']
      emails.append(temp_node.metadata)
    else:
      del temp_node.metadata['DocType']
      del temp_node.metadata['EventType']
      events.append(temp_node.metadata)
        
  email = tabulate(emails, tablefmt="github", headers="keys")
  event = tabulate(events, tablefmt="github", headers="keys")
  md = f"""
  ## Sources
  ### Emails \n
  {email}
  ### Events \n
  {event}
  """
  display(Markdown(md))

def get_text_from_html(html_content):
  soup = BeautifulSoup(html_content, 'html.parser')
  return soup.get_text()

def _get_credentials():
    """Get valid user credentials from storage.

    The file token.json stores the user's access and refresh tokens, and is
    created automatically when the authorization flow completes for the first
    time.

    Returns:
        Credentials, the obtained credential.
    """


    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=8080)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    return creds


def load_cal(
    number_of_results: Optional[int] = 100,
    start_date: Optional[Union[str, datetime.date]] = None,
    end_date: Optional[Union[str, datetime.date]] = None,
) -> List[Document]:
    """Load data from user's calendar.
    Args:
        number_of_results (Optional[int]): the number of events to return. Defaults to 100.
        start_date (Optional[Union[str, datetime.date]]): the start date to return events from. Defaults to today.
    """
    self = GoogleCalendarReader()
    credentials = _get_credentials()
    service = build("calendar", "v3", credentials=credentials)

    start_datetime_utc = start_date.strftime("%Y-%m-%dT%H:%M:%S.%fZ")
    end_datetime_utc = end_date.strftime("%Y-%m-%dT%H:%M:%S.%fZ")

    events_result = (
        service.events()
        .list(
            calendarId="primary",
            timeMin=start_datetime_utc,
            timeMax=end_datetime_utc,
            maxResults=number_of_results,
            singleEvents=True,
            orderBy="startTime",
            eventTypes=["default"]
        )
        .execute()
    )
    # TODO: add pagination
    events = events_result.get("items", [])

    if not events:
        return []

    results = []
    for event in events:
        if event.get("eventType") == "workingLocation":
            continue
        organizer = event.get("organizer", {"email": "N/A", "displayName": "N/A"})

        event_summary = {}
        event_summary['DocType'] = "event"
        event_summary['EventType'] = event['eventType']
        event_summary['id'] = event['id']
        event_summary['url'] = event['htmlLink']
        event_summary['Status'] = event['status']
        event_summary['Start'] = event["start"]["dateTime"] if "dateTime" in event["start"] else event["start"]["date"]
        event_summary['End'] = event["end"]["dateTime"] if "dateTime" in event["end"] else event["end"]["date"]
        event_summary['Summary'] = event['summary']
        if organizer.get('displayName'):
            event_summary['Organizer'] = f"{organizer.get("displayName")} ({organizer['email']})"
        else:
            event_summary['Organizer'] = organizer['email']

        attendees = []
        for attendee in event.get("attendees", []):
            attendees.append({
                "Name": attendee.get("displayName", "N/A"),
                "Email": attendee.get("email", "N/A"),
                "Response Status": attendee.get("responseStatus", "N/A"),
            })

        event_content = []
        # TODO: do something smaller tabulate

        # event_content.append(tabulate([event_summary], headers="keys"))
        event_content.append(json.dumps(event_summary))
        event_content.append(json.dumps(attendees))
        # event_content.append(tabulate(attendees, headers="keys"))
        event_content.append(get_text_from_html(event.get("description", "")))
        event_string = "\n".join(event_content)
        results.append(Document(text=event_string, metadata=event_summary, id_=event['id']))

    return results




def load_email(
    number_of_results: Optional[int] = 100,
    start_date: Optional[Union[str, datetime.date]] = None,
) -> List[Document]:

  query = get_email_query(start_date)
  loader = GmailReader(query=query, results_per_page=100, service=None, max_results=number_of_results)
  credentials=_get_credentials()
  loader.service = build("gmail", "v1", credentials=credentials)

  emails = []
  try:
      messages = loader.search_messages()
  except Exception as e:
      print(e)

  for message in messages:
      try:
        raw_text = message.pop("body")
        raw_email = message_from_string(raw_text)
        
        text = get_plain_text_from_email(raw_email)
        metadata = {
            'id': message.get('id'),
            'threadId': message.get('threadId'),
            'DocType': 'email',
            'url': f"https://mail.google.com/mail/u/0/#inbox/{message.get('id')}",
            'Subject': clean_string(raw_email.get('Subject')),
            'From': clean_string(raw_email.get('From')),
            'To': clean_string(raw_email.get('To')),
            'Cc': clean_string(raw_email.get('Cc')),
            'Bcc': clean_string(raw_email.get('Bcc')),
            'Date': clean_string(raw_email.get('Date'))
        }
        emails.append(Document(text=text, metadata=metadata, id_=message.get('id')))
      except Exception as e:
         print(e)
  return emails


def get_email_query(start_date): 
  filters = [
    'after:' + start_date.strftime('%Y/%m/%d'),
  ]
  negative_subjects = [
    'Appointment booked',
    'invitation',
    'Accepted',
    'Cancelled',
    'document shared',
    'Declined'
  ]

  query_pos = " ".join(filters)
  query_neg = " ".join(f'-Subject:"{subject}"' for subject in negative_subjects)
  return query_pos + " " + query_neg


def get_plain_text_from_email(raw_email, max_length=10000):
    # Parse the raw email string into an EmailMessage object
    message = raw_email



    def decode_payload(payload, charset='utf-8'):
        try:
            return payload.decode(charset)
        except UnicodeDecodeError:
            # Try decoding with a different charset or use 'replace' strategy
            return payload.decode(charset, errors='replace')

    plain_text_parts = []

    if message.is_multipart():
        for part in message.walk():
            content_type = part.get_content_type()
            content_disposition = part.get("Content-Disposition")

            if content_type == 'text/plain' and not content_disposition:
                # Decode the plain text part and add to the list
                payload = part.get_payload(decode=True)
                plain_text_parts.append(decode_payload(payload))
            elif content_type == 'text/html' and not content_disposition:
                # Decode the HTML part, convert to plain text, and add to the list
                payload = part.get_payload(decode=True)
                html_text = decode_payload(payload)
                plain_text_parts.append(get_text_from_html(html_text))
    else:
        # Single part message, could be either text/plain or text/html
        payload = message.get_payload(decode=True)
        content_type = message.get_content_type()
        if content_type == 'text/plain':
            plain_text_parts.append(decode_payload(payload))
        elif content_type == 'text/html':
            html_text = decode_payload(payload)
            plain_text_parts.append(get_text_from_html(html_text))

    # Combine all plain text parts (if multipart message)
    return ('\n'.join(plain_text_parts))[:max_length]

def clean_string(s):
    if (s is None):
        return None
    return s.replace('\r', '').replace('\n', '').replace('\t', '')

def retry_function(func, *args, max_attempts=5, **kwargs):
    attempt = 0
    while attempt < max_attempts:
        try:
            return func(*args, **kwargs)  # Call the function with arguments
        except Exception as e:
            print(f"Attempt {attempt+1} failed with error: {e}")
            attempt += 1
            if attempt == max_attempts:
                raise  # Re-raise the last exception if out of attempts
            # Optionally, log the exception or wait before retrying

