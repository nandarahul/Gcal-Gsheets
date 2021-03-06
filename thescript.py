from __future__ import print_function
import httplib2
import os, sys

import argparse
from apiclient import discovery
import oauth2client

import datetime, calendar

CALENDAR_ID = 'ck12.org_5oalfcv57o5vq0f9dhfsgmr8q0@group.calendar.google.com'

def get_credentials():
    """Gets valid user credentials from storage.
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
    return credentials

def main(month, year):
    """
    """
    last_day_of_month = calendar.monthrange(year,month)[1]
    print(last_day_of_month)

    credentials = get_credentials()
    if not credentials or credentials.invalid:
        print('Please run getCredentials.py to authorise yourself')
        return
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)

    eventsResult = service.events().list(
        calendarId=CALENDAR_ID, timeMin = datetime.datetime(year,month,1).isoformat()+'Z',
        timeMax = datetime.datetime(year,month,last_day_of_month).isoformat()+'Z', 
        singleEvents=True).execute()
    events = eventsResult.get('items', [])
    
    if not events:
        print('No events found.')

    event_dict = {}
    for event in events:
        creator_email = event['creator']['email']
        start = event['start'].get('dateTime', event['start'].get('date'))
        end = event['end'].get('dateTime', event['end'].get('date'))

        start, end = start[:10], end[:10]
        start_date = datetime.datetime.strptime(start, '%Y-%m-%d')    
        end_date = datetime.datetime.strptime(end, '%Y-%m-%d')    

        if 'dateTime' in event['end']:
            end_date += datetime.timedelta(days=1) #end_date is not inclusive according to our logic
        if creator_email not in event_dict:
            event_dict[creator_email] = []

        event_dict[creator_email].append((start_date, end_date))

    print(event_dict)

    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        try:
            print(start, event['summary'], event['creator']['email'])
        except:
            print(start, event['creator']['email'])
    
    #Start writing to google spreadsheet
    import gspread
    gc = gspread.authorize(credentials)
    sh = gc.open("Attendance")
    worksheet_list = sh.worksheets()
    print (worksheet_list)
    sheet_name = str(month) + '/' + str(year)
    #Logic for naming the worksheet
    max = 0
    for sheet in worksheet_list:
        if sheet.title.startswith(sheet_name):
            version = int(sheet.title[-1])
            if version > max:
                max = version
    sheet_name += "-v" + repr(max+1)
    worksheet = sh.add_worksheet(sheet_name, 100, 100)
    for i in range(last_day_of_month):
        worksheet.update_cell(i+2, 1, str(i+1)+'/'+str(month)+'/'+str(year))
    #worksheet = sh.get_worksheet(0)
  
    column = 2
    for creator_email in event_dict.keys():
        worksheet.update_cell(1, column, creator_email)
        for start_end_tuple in event_dict[creator_email]:
            if start_end_tuple[0].month != month: #event started previous month
                day = 1
            else:
                day = start_end_tuple[0].day
            if start_end_tuple[1].month != month: #event ends next month
                day_max = last_day_of_month
            else:
                day_max = start_end_tuple[1].day -1
            while day <= day_max :
                weekday = calendar.weekday(year, month, day)
                if weekday != 5 and weekday != 6:
                    worksheet.update_cell(day+1, column, 'PTO')
                day += 1
        column += 1

if __name__ == '__main__':
    now = datetime.datetime.now()
    parser = argparse.ArgumentParser()
    parser.add_argument(
        '-m', '--month', type=int, dest='month', default=now.month,
        help='The month for which you want PTO info of ck12-in employees'
    )
    parser.add_argument(
        '-y', '--year', type=int, dest='year', default=now.year,
        help='Year'
    )
    args = parser.parse_args()
    #if args:
     #   parser.error('Expected no arguments. Got %r' % args)
    month = args.month
    year = args.year
    print(month, year) 
    main(month, year)
