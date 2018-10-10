from __future__ import print_function
import datetime
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import xlrd

# If modifying these scopes, delete the file token.json.
SCOPES = 'https://www.googleapis.com/auth/calendar'

def main():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('calendar', 'v3', http=creds.authorize(Http()))

    # Load from XLS
    book = xlrd.open_workbook('imi2018.xls')
    curr_sheet = book.sheets()[8]

    _curr_day = ['MO','TU','WE','TH','FR','SA']
    day_counter = -1
    _time_start = ['08:00', '09:50', '11:40', '14:00', '15:45', '17:30']
    _time_end = ['09:35', '11:25', '13:15', '15:35', '17:20', '19:05']

    for i in range(3,39): 
        if (i-3)%6==0:
            day_counter = day_counter+1
        if curr_sheet.cell(i,8).value== "":
            continue
        _name = curr_sheet.cell(i, 8).value
        _type = curr_sheet.cell(i, 9).value
        _location = curr_sheet.cell(i,10).value
        
        event = {
            'summary': _name,
            'location': _location,
            'description': _type,
            'start': {
                'dateTime': '2018-10-08T' + _time_start[(i-3)%6] + ':00+09:00',
                'timeZone': 'Asia/Yakutsk',
            },
            'end': {
                'dateTime':  '2018-10-08T' + _time_end[(i-3)%6] + ':00+09:00',
                'timeZone': 'Asia/Yakutsk',
            },
            'recurrence': [
                'RRULE:FREQ=WEEKLY;BYDAY=' + _curr_day[day_counter]
            ]
        }
        recurring_event = service.events().insert(calendarId='primary', body=event).execute()


if __name__ == '__main__':
    main()