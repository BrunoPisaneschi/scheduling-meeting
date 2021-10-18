from __future__ import print_function
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pathlib
import pandas as pd
from datetime import timedelta

SCOPES = ['https://www.googleapis.com/auth/calendar']


class google_api_calendar():

    def __init__(self, data):
        self._data = data

    def create_invite(self):

        print('creating service...')
        _service = self._create_token_connection_calendar()

        print('creating event...')
        _data_event = self._struct_data()

        event = _service.events().insert(calendarId='primary', body=_data_event, sendUpdates='all').execute()
        print('Event created: %s' % (event.get('htmlLink')))

        return event.get('htmlLink')

    def _create_token_connection_calendar(self):
        creds = None
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        service = build('calendar', 'v3', credentials=creds)

        return service

    def _struct_data(self):

        email_entrevistadores = [
            {'email': self._data['email_responsavel']},
            {'email': self._data['email_candidato']}
        ]

        d_str = str(self._data['data_entrevista']).replace('/', '-') + 'T' + str(self._data['horario_entrevista'])
        date_time_obj = datetime.datetime.strptime(d_str, '%d-%m-%YT%H:%M')

        if (self._data['status_card'] == 'Entrevista Cliente'): email_entrevistadores.append(
            {'email': self._data['email_cliente']})

        event = {
            'summary': ' - ' + self._data['status_card'],
            'location': 'hangouts',
            'description': '',
            'start': {
                'dateTime': date_time_obj.strftime('%Y-%m-%dT%H:%M:%S'),
                'timeZone': 'America/Sao_Paulo',
            },
            'end': {
                'dateTime': (date_time_obj + timedelta(hours=1)).strftime('%Y-%m-%dT%H:%M:%S'),
                'timeZone': 'America/Sao_Paulo',
            },
            'attendees': email_entrevistadores,
            'reminders': {
                'useDefault': False,
                'overrides': [
                    {'method': 'email', 'minutes': 24 * 60},
                    {'method': 'popup', 'minutes': 10},
                ],
            },
        }

        return event
