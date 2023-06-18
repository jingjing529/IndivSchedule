import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from datetime import datetime, timedelta
import os

# If modifying these scopes, delete the file token1.pickle.

def get_calendar_service():
   creds = None
   SCOPES = ['https://www.googleapis.com/auth/calendar']
   CREDENTIALS_FILE = 'credentials.json'
   if os.path.exists('token1.pickle'):
       with open('token1.pickle', 'rb') as token:
           creds = pickle.load(token)
   # If there are no (valid) credentials available, let the user log in.
   if not creds or not creds.valid:
       if creds and creds.expired and creds.refresh_token:
           creds.refresh(Request())
       else:
           flow = InstalledAppFlow.from_client_secrets_file(
               CREDENTIALS_FILE, SCOPES)
           creds = flow.run_local_server(port=0)

       # Save the credentials for the next run
       with open('token1.pickle', 'wb') as token:
           pickle.dump(creds, token)

   service = build('calendar', 'v3', credentials=creds)
   return service


def Event():
    # creates one hour event tomorrow 10 AM IST
    service = get_calendar_service()
    sheet_reader = get_sheet_service()
    for row in sheet_reader:
        date = row[1].split("/") #split the month and day
        TimeHM = row[7][:-1].split(":") #split the hour and minute
        time = datetime(2023, int(date[0]), int(date[1]), int(TimeHM[0]))+timedelta(minutes=int(TimeHM[1]))
        if row[7][-1] == "p" and TimeHM[0] != "12":
            time += timedelta(hours=12)
        start = time.isoformat()
        EndTimeHM = row[8][:-1].split(":") #split the end time hour and minute
        endtime = datetime(2023, int(date[0]), int(date[1]), int(EndTimeHM[0]))+timedelta(minutes=int(EndTimeHM[1]))
        if row[8][-1] == "p" and EndTimeHM[0] != "12":
            endtime += timedelta(hours=12)
        end = endtime.isoformat()
            
        if len(row) > 13: #if there are viewers
            viewer = ", ".join(row[13:])
        else:
            viewer = ""
        event_result = service.events().insert(calendarId='uci.edu_nktaoooi1i2cmq2grcg6thsfck@group.calendar.google.com',
            body={
                "summary": f"[IndivRobotics]sub-{row[6]}_S{row[2]}",
                "description": f'Subject: {row[5]}\nLifeguard Experimenter: {row[9]}\nTechie Experimenter: {row[11]}\nViewer(s): {viewer}\n\nIf you are the pilot subject, please arrive 15 minutes after the scheduled time in the calendar and expect your session to be ~2 hours.\n\nIf you are an experimenter or viewer, please arrive on time as listed in this calendar event. You are expected to be present throughout the full 2.5 hours unless otherwise noted.\n\nDescription of roles:\nLifeguard Experimenter = this experimenter reads instructions during the CAVERN portion and devotes their full attention to the participant in CAVERN\nTechie Experimenter = this experimenter sets up all tasks in both the testing room and CAVERN and reads instructions during the testing room portion\nViewer = this individual is fully attentive to situations that arise during the session, takes detailed notes, and helps wherever needed',
                "start": {"dateTime": start, "timeZone": 'America/Los_Angeles'},
                "end": {"dateTime": end, "timeZone": 'America/Los_Angeles'},
            }
        ).execute()
        
        if row[2] in ["1", "2","3","4","5"]:
            if row[2] == "1": #if it's session1, CAVERN reserves 2 hours
                CAVERNendtime = time + timedelta(hours=2)
            elif row[2] in ["2","3","5"]: #if it's session 2,3,5, CAVERN reserves 1.25 hours
                CAVERNendtime = time + timedelta(hours=1.25)
            else: 
                CAVERNendtime = time + timedelta(hours=1.5)
                #if it's session 4, CAVERN reserves 1.5 hours
            end = CAVERNendtime.isoformat()
            Leads = ["Alina", "Nikki", "Marjan", "Taylor", "Nick", "Olivia", "Volker", "Wesley", "Quinn"]
            if row[9] in Leads:
                access = row[9]
            elif row[11] in Leads:
                access = row[11]
            elif len(row) > 13 and row[13] in Leads:
                access = row[13]
            else:
                access = "N/A"

            CAVERN_result = service.events().insert(calendarId='c_jvkke282k16vkvfj61cgvg1bgc@group.calendar.google.com',
                        body={
                            "summary": '[IndivRobotics] Pilot Participant',
                            "description": f"Unlock CAVERN: {access}\n\nLock CAVERN: {access}",
                            "start": {"dateTime": start, "timeZone": 'America/Los_Angeles'},
                            "end": {"dateTime": end, "timeZone": 'America/Los_Angeles'},
                        }
                    ).execute()


# If modifying these scopes, delete the file token.pickle.
def get_sheet_service():

    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    # here enter the id of your google sheet
    SAMPLE_SPREADSHEET_ID_input = '1xZgC-h5alm6EzZJuzkybhRCZPdellI8aImrrC_oPPdk'
    SAMPLE_RANGE_NAME = 'B3:S14'

    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES) # here enter the name of your downloaded JSON file
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result_input = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID_input,
                                range=SAMPLE_RANGE_NAME).execute()
    values_input = result_input.get('values', [])
    return values_input

if __name__ == "__main__":
    Event()
