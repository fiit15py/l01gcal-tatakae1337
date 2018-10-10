from __future__ import print_function
import datetime
import xlrd
from oauth2client import file, client, tools
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools

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

    calendarStringId = 'rdudthvqgpqjo13t729dqe5b2s@group.calendar.google.com'
    service = build('calendar', 'v3', http=creds.authorize(Http()))
    #service.calendar(calendarStringId).clear().exec


    # Call the Calendar API
    now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time

    group_name = 'ФИИТ-15'
    book = xlrd.open_workbook('imi2018.xls')
    curr_sheet = book.sheets()[8]
    for i in range(20):
        if group_name in curr_sheet.cell(2,i).value: col = i
    for j in range(3, 38):
        dayIndex = 1 + (j - 3) // 6
        time = str(curr_sheet.cell(j,1).value).replace(' -- ', '-').replace('.', ':').strip()
        subj = str(curr_sheet.cell(j,col).value)
        if '-' in time and len(subj) > 1:
            startTime, endTime = time.split('-')
            description = str(curr_sheet.cell(j,col+1).value)
            location = str(curr_sheet.cell(j,col+2).value)
            if (7 + dayIndex) <= 9:
                dayString = '0' + str(7 + dayIndex)
            else:
                dayString = str(7 + dayIndex)
            print(dayString + " :: " + startTime + '/' + endTime + ': ' + subj + ' / ' + description + ' -> ' + location)
            event = {
              'summary': subj,
              'location': location,
              'description': description,
              'start': {
                'dateTime': '2018-10-' + dayString + 'T' + startTime + ':00+09:00',
                'timeZone': 'Asia/Yakutsk',
              },
              'end': {
                'dateTime': '2018-10-' + dayString + 'T' + endTime + ':00+09:00',
                'timeZone': 'Asia/Yakutsk',
              },
              'recurrence': [
                'RRULE:FREQ=WEEKLY;COUNT=12'
              ],
              'reminders': {
              }
            }
            event = service.events().insert(calendarId=calendarStringId, body=event).execute()
            print('Event created: %s' % (event.get('htmlLink')))

#calendarId='primary'

if __name__ == '__main__':
    main()
