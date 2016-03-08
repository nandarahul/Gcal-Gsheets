from __future__ import print_function
import httplib2
import os, sys

#import argparse
from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools

import datetime, calendar

try:
    import argparse
    #flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

SCOPES = ['https://www.googleapis.com/auth/calendar.readonly', 'https://spreadsheets.google.com/feeds']
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'myclient1'
CALENDAR_ID = 'ck12.org_5oalfcv57o5vq0f9dhfsgmr8q0@group.calendar.google.com'
SHEET_NAME = 'Attendance'

if __name__ == '__main__':
    """
    If no credential has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'calendar-python-quickstart.json')
    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    else:
        print('Credentials already exist. You can go ahead and run the script')
