from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import datetime
import os.path
import pickle
import time

import pandas as pd

SCOPES = ["https://www.googleapis.com/auth/calendar"]

def checkDutyStatus(row, weekend):
    status = row % 5
    if status == 4:
        return "Primary"
    elif weekend:
        return "Secondary"
    elif status == 0:
        return "Secondary"
    elif status == 1:
        return "Desk"
    elif status == 2:
        return "Off"

def checkWhoWorkingWith(row, weekend, dayColum):
    if row % 5 == 2: # If we are primary
        primary = dayColum[row]
        secondary1 = dayColum[row + 1]
        secondary2 = dayColum[row + 2]
        secondary3 = dayColum[row + 3]
    elif row % 5 == 3: # If we are 1st under primary
        primary = dayColum[row - 1]
        secondary1 = dayColum[row]
        secondary2 = dayColum[row + 1]
        secondary3 = dayColum[row + 2]
    elif row % 5 == 4: # If we are 2nd under primary
        primary = dayColum[row - 2]
        secondary1 = dayColum[row - 1]
        secondary2 = dayColum[row]
        secondary3 = dayColum[row + 1]
    else: # If we are 3rd under primary
        primary = dayColum[row - 3]
        secondary1 = dayColum[row - 2]
        secondary2 = dayColum[row - 1]
        secondary3 = dayColum[row]
    if weekend == 1:
        return {
                "Primary": primary,
                "1st Secondary": secondary1,
                "2nd Secondary": secondary2,
                "3rd Secondary": secondary3,
        }
    else:
        return {
            "Primary": primary,
            "Secondary": secondary1,
            "Desk": secondary2,
            "Off": secondary3
        }

def checkWeekendDuty(row, day, satDayDuty, sunDayDuty):
    saturday1 = "N/A"
    saturday2 = "N/A"
    sunday1 = "N/A"
    sunday2 = "N/A"
    if row % 5 == 2: # If we are primary
        saturday1 = satDayDuty[row]
        saturday2 = satDayDuty[row + 1]
        sunday1 = sunDayDuty[row]
        sunday2 = sunDayDuty[row + 1]
    elif row % 5 == 3:
        saturday1 = satDayDuty[row - 1]
        saturday2 = satDayDuty[row]
        sunday1 = sunDayDuty[row - 1]
        sunday2 = sunDayDuty[row]
    elif row % 5 == 4:
        saturday1 = satDayDuty[row - 2]
        saturday2 = satDayDuty[row - 1]
        sunday1 = sunDayDuty[row - 2]
        sunday2 = sunDayDuty[row - 1]
    else:
        saturday1 = satDayDuty[row - 3]
        saturday2 = satDayDuty[row - 2]
        sunday1 = sunDayDuty[row - 3]
        sunday2 = sunDayDuty[row - 2]

    return{
        "Saturday 1":saturday1,
        "Saturday 2":saturday2,
        "Sunday 1":sunday1,
        "Sunday 2":sunday2
    }

def dayStart(row, dayColum):
    if row % 5 == 2: # If we are primary
        startDay = dayColum[row-1]
    elif row % 5 == 3:
        startDay = dayColum[row-2]
    elif row % 5 == 4:
        startDay = dayColum[row-3]
    else:
        startDay = dayColum[row-4]
    nextDay = startDay + 1
    startDay = f"{startDay:02d}"
    nextDay = f"{nextDay:02d}"
    return startDay, nextDay

def monthToNum(month_name):
    sheetNames = ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY"]
    month_num = {month: idx + 1 for idx, month in enumerate(sheetNames)}

    if month_name in month_num:
        return month_num[month_name]
    else:
        raise ValueError(f"Invalid month name: {month_name}")

def formatDescription(coworkers, weekendDuty):
    """
    Formats the coworkers and weekend duty dictionaries into a neat string.

    Args:
        coworkers (dict): Dictionary with roles as keys and names as values.
        weekendDuty (dict, optional): Dictionary with weekend duty roles as keys and names as values.

    Returns:
        str: A formatted string representation of the coworkers and weekend duty.
    """
    formatted_description = "Coworkers:\n"
    
    # Format coworkers
    for role, name in coworkers.items():
        formatted_description += f"- {role}: {name.strip() if isinstance(name, str) else 'N/A'}\n"
    
    if len(weekendDuty) > 0:
        # Add a separator before weekend duty
        formatted_description += "\nWeekend Duty:\n"
        
        # Format weekend duty
        for day, name in weekendDuty.items():
            formatted_description += f"- {day}: {name.strip() if isinstance(name, str) else 'N/A'}\n"

    return formatted_description.strip()  # Remove trailing newline

def main():
    creds = None

    sheetNames = ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY"]
    days = ["SUNDAY", "MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY"]
    name = "Ian"

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
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        # Build the Calendar API service
        service = build("calendar", "v3", credentials=creds)
        
        # Define the calendar properties
        #calTitle = f"{name}s Duty Schedule"
        calendar = {
            'summary': (f'{name}\'s Duty Schedule'),  # Replace with your desired calendar title
            'timeZone': 'US/Eastern'
        }
        
        # Create the new calendar
        created_calendar = service.calendars().insert(body=calendar).execute()
        
        # Get the calendar ID
        calendar_id = created_calendar['id']
        
        # Generate the public link
        public_link = f"https://calendar.google.com/calendar/embed?src={calendar_id}"
        
        print(f"Calendar created successfully!")
        print(f"Calendar Link: {public_link}")

        for sheet in sheetNames:
            time.sleep(0.25)
            df = pd.read_excel("Final MCUT Duty 24-25.xlsx", sheet_name=sheet)
            for day in range(7):
                weekend = 0
                if day >= 4:
                    weekend = 1
                    satDayDuty = df.iloc[:, 7]
                    sunDayDuty = df.iloc[:, 8]
                dayColum = df.iloc[:, day]
                
                for idx, value in enumerate(dayColum):
                    dutyStatus = checkDutyStatus(idx + 2, weekend)
                    if value == name:
                        coworkers = checkWhoWorkingWith(idx,weekend, dayColum)
                        start = dayStart(idx, dayColum)
                        numMonth = monthToNum(sheet)
                        weekendDuty = ""
                        if weekend:
                            weekendDuty = checkWeekendDuty(idx, day, satDayDuty, sunDayDuty)
                        formattedDescription = formatDescription(coworkers, weekendDuty)
                        if dutyStatus == "Desk":
                            timeStart = (f'2025-{numMonth}-{start[0]}T20:00:00-05:00')
                            timeEnd = (f'2025-{numMonth}-{start[0]}T23:59:59-05:00')
                        else:
                            timeStart = (f'2025-{numMonth}-{start[0]}T19:00:00-05:00')
                            timeEnd = (f'2025-{numMonth}-{start[1]}T08:00:00-05:00')
                        event = {
                            'summary': dutyStatus,
                            'description': formattedDescription,
                            'start': {
                                'dateTime': timeStart,
                                'timeZone': 'US/Eastern',
                            },
                            'end': {
                                'dateTime': timeEnd,
                                'timeZone': 'US/Eastern',
                            },
                            'reminders': {
                                'useDefault': False,
                                'overrides': [
                                    {'method': 'popup', 'minutes': 2 * 60},
                                    {'method': 'popup', 'minutes': 30},
                                ],
                            },
                        }
                        event = service.events().insert(calendarId=calendar_id, body=event).execute()
                        #print(f"\nFound '{name}' in sheet '{sheet}' on day {days[day]} at row {idx + 2} and is {dutyStatus}")
                        #print(f"Coworkers: {coworkers}")
                        #print(f"Weekend Duty Workers: {weekendDuty}")
        
    except HttpError as error:
        print(f"An error occurred: {error}")
        print(f"Details: {error.error_details}")

if __name__ == "__main__":
    main()
