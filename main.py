import os
import time
import copy
from pprint import pprint
import re
from datetime import datetime

import google_auth_oauthlib.flow
import googleapiclient.discovery
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request

class Google_Calendar(object):

  def __init__(self):

    # Make sure you have those files/workspace ready
    self.cloud_workspace = '/content/drive/MyDrive/Colab Notebooks'
    self.timetable_path = f'{self.cloud_workspace}/timetable.rtf'
    # Google Cloud OAuth 2.0 Client IDs credential
    # https://developers.google.com/identity/protocols/oauth2
    self.credentails_path = f'{self.cloud_workspace}/google_calendar.json'
    
    self.cred = f'{self.cloud_workspace}/cred.json'
    self.credentials = self.get_credentials()
    self.service = googleapiclient.discovery.build('calendar', 'v3', credentials=self.credentials)

  def get_credentials(self):

    SCOPES = ['https://www.googleapis.com/auth/calendar']
    
    if os.path.exists(self.cred):
      credentials = Credentials.from_authorized_user_file(self.cred, SCOPES)
      if credentials and credentials.expired and credentials.refresh_token:
        credentials.refresh(Request())
    else:
      flow = google_auth_oauthlib.flow.Flow.from_client_secrets_file(
        client_secrets_file=self.credentails_path,
        scopes=SCOPES,
        redirect_uri='https://localhost'
      )
      auth_url = flow.authorization_url()
      print('Go to this URL to authorize to get code: {}'.format(auth_url))
      code = input('Enter the authorization code: ')
      flow.fetch_token(code=code)
      credentials=flow.credentials
      with open(self.cred, 'w+') as f:
        f.write(flow.credentials.to_json())
    return credentials

  def trigger_event(self, event):
    result = self.service.events().insert(calendarId='primary', body=event).execute()
    print('Event created: {}'.format(result.get('id')))

  def trim_Tstring(self, x):
    date_obj = datetime.strptime(x.strip(),'%d-%b-%Y')
    return date_obj.strftime('%Y-%m-%d')

  def count_weeks(self, start, end):
    start_date_obj = datetime.strptime(start.strip(),'%d-%b-%Y')
    end_date_obj = datetime.strptime(end.strip(),'%d-%b-%Y')
    return (end_date_obj - start_date_obj).days // 7

  def read_time(self, string):
      blocks = string.split('\n\n')
      arr = [dict([line.strip().split(': ', maxsplit=1) for line in lines.strip().split('\n')]) for lines in blocks]
      for i in arr:
        start, end = i['Time'].split(' until ')
        i['recurrence'] = self.count_weeks(i['Start Date'], i['End Date'])
        i['End Time'] = f"{self.trim_Tstring(i['Start Date'])}T{end}:00"
        i['Start Date'] = f"{self.trim_Tstring(i['Start Date'])}T{start}:00"
      return arr
  
  def create_biweekly_event_block(self, d):
    pprint(d)
    return {
          'summary': f"{d['Course Name']} {d['Course Code']} {d['Section']}",
          'location': d.get('Room Number/ Location'),
          'description': "\n".join([f"{k}: {v}" for k,v in x.items() if k not in ["End Time","","recurrence"]]),
          'start': {
              'dateTime': d['Start Date'],
              'timeZone': 'America/Toronto',
          },
          'end': {
              'dateTime': d['End Time'],
              'timeZone': 'America/Toronto',
          },
          'recurrence': [
              f'RRULE:FREQ=WEEKLY;COUNT={d["recurrence"]}'
          ],
          'reminders': {
              'useDefault': False,
              'overrides': [
                  {'method': 'popup', 'minutes': 30},
              ],
          },
      }
    
  def process(self):
    with open(self.timetable_path, 'r') as f:
      table = f.read()
    data = self.read_time(table)
    for i in data:
      event_block = self.create_biweekly_event_block(i)
      self.trigger_event(event_block)
      time.sleep(0.1)
      break


g = Google_Calendar()
g.process()