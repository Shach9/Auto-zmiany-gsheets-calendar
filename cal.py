from __future__ import print_function

import datetime
import os.path
import gspread

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


SCOPES = ['https://www.googleapis.com/auth/calendar', 'https://www.googleapis.com/auth/calendar.events']
key='AIzaSyDdNRstvFVVyWjXKsZi1n-tVL3i7z1Em84'   


def main():
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

    try:
        service = build('calendar', 'v3', credentials=creds)
        now = datetime.datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
        
        #Acessesing sheet
        sa = gspread.service_account()
        sh = sa.open("Cinema")
        wks = sh.worksheet("Godziny")



        #Creating events
        row = 4
        while row < 24:
            #Getting shift type and shift time values
            row = str(row)
            dzien = wks.acell('C' + row).value
            data_start = wks.acell('D'+ row).value
            data_end = wks.acell('E' + row).value
            zmiana = wks.acell('B' + row).value
            if data_start == None or data_end == None or dzien == None or zmiana == None:
                print('No values left, exiting program')
                break            


            #Converting date and time to datetime
            date = dzien
            start_time = data_start
            end_time = data_end
            start_date_str = date+start_time
            end_date_str = date + end_time

            start_date = str(datetime.datetime.strptime(start_date_str, '%Y-%m-%d%H:%M').isoformat())
            end_date = str(datetime.datetime.strptime(end_date_str, '%Y-%m-%d%H:%M').isoformat())
        

            if zmiana == 'OW':
            
                event = {
                'summary': 'OW',
                'colorId': '6',
                'start': {
                    'dateTime': start_date,
                    'timeZone': 'Europe/Warsaw',
                },
                'end': {
                    'dateTime': end_date,
                    'timeZone': 'Europe/Warsaw',
                },
                'reminders': {
                    'useDefault': False,
                    'overrides': [
                    {'method': 'popup', 'minutes': 5},
                    ],
                },
                }
            elif zmiana == 'BAR/OW':
            
                event = {
                'summary': 'BAR/OW',
                'colorId': '6',
                'start': {
                    'dateTime': start_date,
                    'timeZone': 'Europe/Warsaw',
                },
                'end': {
                    'dateTime': end_date,
                    'timeZone': 'Europe/Warsaw',
                },
                'reminders': {
                    'useDefault': False,
                    'overrides': [
                    {'method': 'popup', 'minutes': 5},
                    ],
                },
                }    
            elif zmiana == 'BAR':
                event = {
                'summary': 'BAR',
                'colorId': '6',
                'start': {
                    'dateTime': start_date,
                    'timeZone': 'Europe/Warsaw',
                },
                'end': {
                    'dateTime': end_date,
                    'timeZone': 'Europe/Warsaw',
                },
                'reminders': {
                    'useDefault': False,
                    'overrides': [
                    {'method': 'popup', 'minutes': 5},
                    ],
                },
                }                    
            else:
                event = {
                'summary': 'NOKTOWIZOR',
                'colorId': '6',
                'start': {
                    'dateTime': start_date,
                    'timeZone': 'Europe/Warsaw',
                },
                'end': {
                    'dateTime': end_date,
                    'timeZone': 'Europe/Warsaw',
                },
                'reminders': {
                    'useDefault': False,
                    'overrides': [
                    {'method': 'popup', 'minutes': 5},
                    ],
                },
                }
            event = service.events().insert(calendarId='primary', body=event).execute()
            print ('Event created: %s' % (event.get('htmlLink')))
            row = int(row)
            row += 1 

    except HttpError as error:
        print('An error occurred: %s' % error)


if __name__ == '__main__':
    main()

