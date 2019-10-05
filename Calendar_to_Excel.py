from googleapiclient.discovery import build
import xlsxwriter
import pickle
import os
from operator import itemgetter
from datetime import datetime
import datetime
import copy
import calendar

# first run on a new user
# from apiclient import discovery
# from google_auth_oauthlib.flow import InstalledAppFlow
# scopes = ['https://www.googleapis.com/auth/calendar']
# flow = InstalledAppFlow.from_client_secrets_file("client_secret1.json", scopes=scopes)
# credentials = flow.run_console()
# pickle.dump(credentials, open("token.pkl", "wb"))

# Load credentials from pickle file.
credentials = pickle.load(open("token.pkl", "rb"))
service = build("calendar", "v3", credentials=credentials)

# resultCalendarList is a list with calendars.
resultCalendarList = service.calendarList().list().execute()

# The first calendar in the list where [0] represents the default calendar and can be replaced.
calendar_id = resultCalendarList["items"][0]["id"]

# Input the start and end dates for Officers' Workshop sheet.(for the data to be gathered and processed)
START_DAY_OF_OFFICERS_WORKSHOP = 19  # Start Day must be a Monday.
START_MONTH_OF_OFFICERS_WORKSHOP = 11
START_YEAR_OF_OFFICERS_WORKSHOP = 2018
END_DAY_OF_OFFICERS_WORKSHOP = 27
END_MONTH_OF_OFFICERS_WORKSHOP = 3
END_YEAR_OF_OFFICERS_WORKSHOP = 2019

# Iterate over the calendar to get a list with all events : resultEventsList.
# For a specific period of time add as argument in service.events().list
# the following: timeMin="2019-07-01T01:03:12Z" ,timeMax="2019-11-05T01:03:12Z". e.g. below commented.
resultEventsList = []
page_token = None
while True:
    events = (
        service.events()
        .list(
            calendarId=calendar_id,
            singleEvents=True,
            orderBy="startTime",
            pageToken=page_token,
            # timeMin="2019-07-01T01:03:12Z"
            # timeMax="2019-11-05T01:03:12Z"
        )
        .execute()
    )
    page_token = events.get("nextPageToken")
    resultEventsList = resultEventsList + events["items"]
    if not page_token:
        break

# Create a workbook and add a worksheets.
if os.path.exists("CalendarT.xlsx"):  # Override the current Calendar File
    os.remove("CalendarT.xlsx")

# Alternative date format for dates: worksheet.set_column('B:B',None,cell_format) instead of just the default one.
workbook = xlsxwriter.Workbook("CalendarT.xlsx", {"default_date_format": "dd/mm/yy"})
worksheetB = workbook.add_worksheet("Workshops Booking")
worksheet = workbook.add_worksheet("Workshop Data")
worksheetE = workbook.add_worksheet("Events")
worksheetT = workbook.add_worksheet("Teacher Events")
worksheetMe = workbook.add_worksheet("Meetings")
worksheetTc = workbook.add_worksheet("Club Events")
worksheetS = workbook.add_worksheet("Special Events")
worksheetM = workbook.add_worksheet("Misc.")
worksheetW = workbook.add_worksheet("Officers' Workshop")

# Raw data.
months = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
]
days = ["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"]

# Format headers.
bold = workbook.add_format({"bold": True})
# Format cells.
alignment = workbook.add_format({"align": "left"})
alignmentC = workbook.add_format({"align": "center"})
# Format dates.
cell_format = workbook.add_format({"num_format": "dd/mm/yy", "align": "left"})

worksheetB.write("A1", "Hub", bold)
worksheetB.write("B1", "Date", bold)
worksheetB.write("C1", "Start", bold)
worksheetB.write("D1", "End", bold)
worksheetB.write("E1", "School Number", bold)
worksheetB.write("F1", "School", bold)
worksheetB.write("G1", "KS", bold)
worksheetB.write("H1", "Workshop", bold)
worksheetB.write("I1", "Number of Pupils", bold)
worksheetB.write("J1", "Attendee 1", bold)
worksheetB.write("K1", "Attendee 2", bold)
worksheetB.write("L1", "Attendee 3", bold)
worksheetB.write("M1", "Attendee 4", bold)
worksheetB.write("N1", "Attendee 5", bold)
worksheetB.write("O1", "Rest of Attendees", bold)
worksheetB.write("P1", "Teacher Contact", bold)
worksheetB.write("Q1", "Workshop Number", bold)
worksheetB.write("R1", "Workshop ID", bold)
worksheetB.write("S1", "Initial Paperwork", bold)
worksheetB.write("T1", "Initial Paperwork returned", bold)
worksheetB.write("U1", "MIF", bold)
worksheetB.write("V1", "IIF", bold)
worksheetB.write("W1", "SA missing", bold)
worksheetB.write("X1", "Year Group", bold)
worksheetB.write("Y1", "Notes", bold)

worksheet.write("A1", "Hub", bold)
worksheet.write("B1", "Date", bold)
worksheet.write("C1", "Start", bold)
worksheet.write("D1", "End", bold)
worksheet.write("E1", "School Number", bold)
worksheet.write("F1", "School", bold)
worksheet.write("G1", "KS", bold)
worksheet.write("H1", "Workshop", bold)
worksheet.write("I1", "Number of Pupils", bold)
worksheet.write("J1", "Teacher Contact", bold)
worksheet.write("K1", "Workshop Number", bold)
worksheet.write("L1", "Workshop ID", bold)
worksheet.write("M1", "Initial Paperwork", bold)
worksheet.write("N1", "Initial Paperwork returned", bold)
worksheet.write("O1", "MIF", bold)
worksheet.write("P1", "IIF", bold)
worksheet.write("Q1", "SA missing", bold)
worksheet.write("R1", "Year Group", bold)
worksheet.write("S1", "Notes", bold)
worksheet.write("T1", "Attendee 1", bold)
worksheet.write("U1", "Attendee 2", bold)
worksheet.write("V1", "Attendee 3", bold)
worksheet.write("W1", "Attendee 4", bold)
worksheet.write("X1", "Attendee 5", bold)
worksheet.write("Y1", "Rest of Attendees", bold)

worksheetE.write("A1", "Hub", bold)
worksheetE.write("B1", "Date", bold)
worksheetE.write("C1", "Start", bold)
worksheetE.write("D1", "End", bold)
worksheetE.write("E1", "Name of Event", bold)
worksheetE.write("F1", "Location", bold)
worksheetE.write("G1", "Attendee 1", bold)
worksheetE.write("H1", "Attendee 2", bold)
worksheetE.write("I1", "Attendee 3", bold)
worksheetE.write("J1", "Attendee 4", bold)
worksheetE.write("K1", "Attendee 5", bold)
worksheetE.write("L1", "Rest of Attendees", bold)
worksheetE.write("M1", "Notes", bold)

worksheetT.write("A1", "Hub", bold)
worksheetT.write("B1", "Date", bold)
worksheetT.write("C1", "Start", bold)
worksheetT.write("D1", "End", bold)
worksheetT.write("E1", "Name of Event", bold)
worksheetT.write("F1", "Location", bold)
worksheetT.write("G1", "Attendee 1", bold)
worksheetT.write("H1", "Attendee 2", bold)
worksheetT.write("I1", "Attendee 3", bold)
worksheetT.write("J1", "Attendee 4", bold)
worksheetT.write("K1", "Attendee 5", bold)
worksheetT.write("L1", "Rest of Attendees", bold)
worksheetT.write("M1", "Number of Teachers", bold)
worksheetT.write("N1", "Notes", bold)

worksheetMe.write("A1", "Hub", bold)
worksheetMe.write("B1", "Date", bold)
worksheetMe.write("C1", "Start", bold)
worksheetMe.write("D1", "End", bold)
worksheetMe.write("E1", "Name of Event", bold)
worksheetMe.write("F1", "Location", bold)
worksheetMe.write("G1", "Attendee 1", bold)
worksheetMe.write("H1", "Attendee 2", bold)
worksheetMe.write("I1", "Attendee 3", bold)
worksheetMe.write("J1", "Attendee 4", bold)
worksheetMe.write("K1", "Attendee 5", bold)
worksheetMe.write("L1", "Rest of Attendees", bold)
worksheetMe.write("M1", "Notes", bold)

worksheetTc.write("A1", "Hub", bold)
worksheetTc.write("B1", "Date", bold)
worksheetTc.write("C1", "Start", bold)
worksheetTc.write("D1", "End", bold)
worksheetTc.write("E1", "Name", bold)
worksheetTc.write("F1", "Number of Pupils", bold)
worksheetTc.write("G1", "Teacher Contact", bold)
worksheetTc.write("H1", "Workshop Number", bold)
worksheetTc.write("I1", "Workshop ID", bold)
worksheetTc.write("J1", "Initial Paperwork", bold)
worksheetTc.write("K1", "Initial Paperwork returned", bold)
worksheetTc.write("L1", "MIF", bold)
worksheetTc.write("M1", "IIF", bold)
worksheetTc.write("N1", "SA missing", bold)
worksheetTc.write("O1", "Year Group", bold)
worksheetTc.write("P1", "Notes", bold)
worksheetTc.write("Q1", "Attendee 1", bold)
worksheetTc.write("R1", "Attendee 2", bold)
worksheetTc.write("S1", "Attendee 3", bold)
worksheetTc.write("T1", "Attendee 4", bold)
worksheetTc.write("U1", "Attendee 5", bold)
worksheetTc.write("V1", "Rest of Attendees", bold)
worksheetTc.write("W1", "Location", bold)

worksheetS.write("A1", "Hub", bold)
worksheetS.write("B1", "Date", bold)
worksheetS.write("C1", "Start", bold)
worksheetS.write("D1", "End", bold)
worksheetS.write("E1", "Name of Event", bold)
worksheetS.write("F1", "Location", bold)
worksheetS.write("G1", "Attendee 1", bold)
worksheetS.write("H1", "Attendee 2", bold)
worksheetS.write("I1", "Attendee 3", bold)
worksheetS.write("J1", "Attendee 4", bold)
worksheetS.write("K1", "Attendee 5", bold)
worksheetS.write("L1", "Rest of Attendees", bold)
worksheetS.write("M1", "Notes", bold)

worksheetM.write("A1", "Date", bold)
worksheetM.write("B1", "Start", bold)
worksheetM.write("C1", "End", bold)
worksheetM.write("D1", "Event Name", bold)
worksheetM.write("E1", "Attendee 1", bold)
worksheetM.write("F1", "Attendee 2", bold)
worksheetM.write("G1", "Attendee 3", bold)
worksheetM.write("H1", "Attendee 4", bold)
worksheetM.write("I1", "Attendee 5", bold)
worksheetM.write("J1", "Rest of Attendees", bold)
worksheetM.write("K1", "Notes", bold)

# Format spacing for columns and date format.
worksheetB.set_column(0, 0, 4.3)
worksheetB.set_column(1, 1, 12)
worksheetB.set_column(2, 3, 8)
worksheetB.set_column(4, 4, 14)
worksheetB.set_column(5, 5, 40)
worksheetB.set_column(7, 7, 40)
worksheetB.set_column(8, 8, 17)
worksheetB.set_column(9, 23, 15)
worksheetB.set_column(14, 16, 17)
worksheetB.set_column(18, 19, 17)
worksheetB.set_column(24, 24, 80)

worksheet.set_column(0, 0, 4.3)
worksheet.set_column(1, 1, 12)
worksheet.set_column(2, 3, 8)
worksheet.set_column(4, 4, 14)
worksheet.set_column(5, 5, 40)
worksheet.set_column(7, 7, 40)
worksheet.set_column(8, 10, 17)
worksheet.set_column(11, 23, 15)
worksheet.set_column(12, 12, 17)
worksheet.set_column(13, 13, 17)
worksheet.set_column(18, 18, 80)
worksheet.set_column(24, 24, 17)

worksheetE.set_column(0, 0, 4.3)
worksheetE.set_column(1, 1, 20)
worksheetE.set_column(2, 3, 8)
worksheetE.set_column(4, 5, 45)
worksheetE.set_column(6, 10, 15)
worksheetE.set_column(11, 11, 17)
worksheetE.set_column(12, 12, 80)

worksheetT.set_column(0, 0, 4.3)
worksheetT.set_column(1, 1, 12)
worksheetT.set_column(2, 3, 8)
worksheetT.set_column(4, 5, 40)
worksheetT.set_column(6, 10, 15)
worksheetT.set_column(11, 12, 18)
worksheetT.set_column(13, 13, 80)

worksheetMe.set_column(0, 0, 4.3)
worksheetMe.set_column(1, 1, 12)
worksheetMe.set_column(2, 3, 8)
worksheetMe.set_column(4, 5, 40)
worksheetMe.set_column(6, 10, 15)
worksheetMe.set_column(11, 11, 17)
worksheetMe.set_column(12, 12, 80)

worksheetTc.set_column(0, 0, 4.3)
worksheetTc.set_column(1, 1, 12)
worksheetTc.set_column(2, 3, 8)
worksheetTc.set_column(4, 4, 40)
worksheetTc.set_column(5, 6, 17)
worksheetTc.set_column(7, 10, 17)
worksheetTc.set_column(11, 20, 15)
worksheetTc.set_column(15, 15, 80)
worksheetTc.set_column(21, 21, 17)
worksheetTc.set_column(22, 22, 40)

worksheetS.set_column(0, 0, 4.3)
worksheetS.set_column(1, 1, 10)
worksheetS.set_column(2, 3, 8)
worksheetS.set_column(4, 5, 45)
worksheetS.set_column(6, 10, 15)
worksheetS.set_column(11, 11, 17)
worksheetS.set_column(12, 12, 80)

worksheetM.set_column(0, 0, 20)
worksheetM.set_column(1, 2, 10)
worksheetM.set_column(3, 3, 40)
worksheetM.set_column(4, 8, 15)
worksheetM.set_column(9, 9, 17)
worksheetM.set_column(10, 10, 80)

# Template for columns.
EVENT_INFORMATION = [
    "Number of Pupils:",
    "Teacher Contact:",
    "Workshop Number:",
    "Workshop ID:",
    "Initial Paperwork:",
    "Initial Paperwork Returned:",
    "MIF:",
    "IIF:",
    "SA Missing:",
    "Year Group:",
    "Number of Teachers:",
    "School Number:",
]

# Hubs.
HUBS = ["HUB1", "HUB2", "HUB3", "HUB4", "HUB5", "HUB6", "HUB7", "HUB8", "HUB9"]

# List with PEOPLE_ATTENDEES's emails and how they will appear in calendar.
PEOPLE_ATTENDEES = {
    "bob.howard@gmail.com": "Bob 1",
    "bobino.howardino@gmail.com": "Bobino",
    "roger.donner@gmail.com": "Roger",
    "george.tan@gmail.com": "George T"
}

# List with PEOPLE WORKSHOPS's emails and how they will appear in Officers' Workshop sheet.
PEOPLE_WORKSHOPS = {
    "bob.howard@gmail.com": "Bob 1",
    "bobino.howardino@gmail.com": "Bobino",
    "roger.donner@gmail.com": "Roger",
    "george.tan@gmail.com": "George T"
}

# List of IGNORED PEOPLE's emails.
IGNORED_PEOPLE = {}

# List of unwanted events.
U_EVENTS = {"HUB1 - Meeting - Team Time"}

# List of special events like summer school.
SPECIAL_EVENTS = {"Summer School"}

# Temporary lists used to build the lists of events.
listA = []  # Workshops Data
listB = []  # Misc
listE = []  # Event
listF = []  # Teacher Events
listG = []  # Meeting
listT = []  # Club Events
listS = []  # Special Events

# Lists of events for each sheet.
listC = []  # Workshops Data
listD = []  # Misc
listH = []  # Event
listI = []  # Teacher Events
listJ = []  # Meeting
listMi = []  # Workshops Booking will be listC with columns reordered.
listTc = []  # Club Events
listSS = []  # Special Events

# Officers' Workshop formatting
i = 0
j = 3
while i <= len(PEOPLE_WORKSHOPS) - 1:
    worksheetW.write(j, 0, list(PEOPLE_WORKSHOPS.values())[i])
    i += 1
    j += 1

# Input the start and end dates for Officers' Workshop sheet .
startD = START_DAY_OF_OFFICERS_WORKSHOP  # Start Day must be a Monday.
startM = START_MONTH_OF_OFFICERS_WORKSHOP
startY = START_YEAR_OF_OFFICERS_WORKSHOP
endD = END_DAY_OF_OFFICERS_WORKSHOP
endM = END_MONTH_OF_OFFICERS_WORKSHOP
endY = END_YEAR_OF_OFFICERS_WORKSHOP

startDate = datetime.datetime.strptime(
    str(startD) + "-" + str(startM) + "-" + str(startY), "%d-%m-%Y"
)
endDate = datetime.datetime.strptime(
    str(endD) + "-" + str(endM) + "-" + str(endY), "%d-%m-%Y"
)
weeksBetweenDates = (
    datetime.date(endY, endM, endD) - datetime.date(startY, startM, startD)
).days / 7
worksheetW.set_column(1, int(weeksBetweenDates * 7), 10)
lineIndex = 0
columnIndex = 1
endW = 0
week = 0
# Format the Officers' Workshop sheet information.
while week < weeksBetweenDates:
    nextMonth = False
    if startD + 6 > calendar.monthrange(startY, startM)[1]:
        nextMonth = True
        endW = startD + 6 - calendar.monthrange(startY, startM)[1]
    else:
        endW = startD + 6
    worksheetW.merge_range(
        lineIndex, columnIndex, lineIndex, columnIndex + 6, months[startM - 1]
    )  # Month title.
    worksheetW.merge_range(
        lineIndex + 1,
        columnIndex,
        lineIndex + 1,
        columnIndex + 6,
        "Week " + str(startD) + " - " + str(endW),
    )  # Week range.
    wj = 0  # Days titles:Mo,Tu,Wen,...
    # Increment to next week.
    while wj + columnIndex < columnIndex + 7:
        worksheetW.write(lineIndex + 2, wj + columnIndex, days[wj])
        wj += 1
    columnIndex += 7  # Increment columns number to next week
    if nextMonth == True or startD + 7 > calendar.monthrange(startY, startM)[1]:
        startD = startD + 7 - calendar.monthrange(startY, startM)[1]
        if startM == 12:
            startY += 1
            startM = 1
        else:
            startM += 1
    else:
        startD += 7  # increment startD to next week
    week += 1


lineIndex = 3
columnIndex = 1
# Introduce dates and times into the list.
def dateB(dateTimeS, dateTimeE, list, type):
    year = dateTimeS[0:4]
    month = dateTimeS[5:7]
    day = dateTimeS[8:10]
    if type == True:
        startTime = dateTimeS[11:16]
        endTime = dateTimeE[11:16]
        list.insert(
            1, datetime.datetime.strptime(day + "-" + month + "-" + year, "%d-%m-%Y")
        )
        list.insert(2, startTime)
        list.insert(3, endTime)
    else:
        yearE = dateTimeE[0:4]
        monthE = dateTimeE[5:7]
        dayE = dateTimeE[8:10]
        list.insert(1, "Multiple all day Event")
        list.insert(
            2, datetime.datetime.strptime(day + "-" + month + "-" + year, "%d-%m-%Y")
        )
        list.insert(
            3, datetime.datetime.strptime(dayE + "-" + monthE + "-" + yearE, "%d-%m-%Y")
        )


# Introduce workshop for each attendee in the Officers' Workshop
def workshops(attendees, listTemp):
    for attendee in attendees:
        if attendee.get("organizer") == None:
            if attendee.get("email") != None:
                if attendee.get("email") not in IGNORED_PEOPLE:
                    name = attendee["email"]
                    status = attendee["responseStatus"]
                    if status == "accepted":
                        # Only PEOPLE_ATTENDEES that confiremed their attendence will appear.
                        if PEOPLE_WORKSHOPS.get(name) != None:
                            if type(listTemp[1]) == datetime.datetime:
                                if listTemp[1] >= startDate and listTemp[1] <= endDate:
                                    daysIndex = (listTemp[1] - startDate).days
                                    worksheetW.write(
                                        lineIndex
                                        + list(PEOPLE_WORKSHOPS.keys()).index(name),
                                        columnIndex + daysIndex,
                                        "Workshop",
                                    )
                            else:
                                if listTemp[2] >= startDate and listTemp[2] <= endDate:
                                    daysIndex = (listTemp[2] - startDate).days
                                    multipleDays = (listTemp[3] - startDate).days - (
                                        listTemp[2] - startDate
                                    ).days
                                    temp = 0
                                    while temp < multipleDays:
                                        worksheetW.write(
                                            lineIndex
                                            + list(PEOPLE_WORKSHOPS.keys()).index(name),
                                            columnIndex + daysIndex + temp,
                                            "Workshop",
                                        )
                                        temp += 1


# Introduce attendees into the calendar.
def attendeesF(attendees, listTemp, column_index):
    i = 0  # attende number
    for attendee in attendees:
        if attendee.get("organizer") == None:
            if attendee.get("email") != None:
                if attendee.get("email") not in IGNORED_PEOPLE:
                    name = attendee["email"]
                    status = attendee["responseStatus"]
                    if status == "accepted":
                        # Only PEOPLE_ATTENDEES that confiremed their attendence will appear.
                        if PEOPLE_ATTENDEES.get(name) != None:
                            if i < 6:
                                listTemp.insert(
                                    column_index + i, PEOPLE_ATTENDEES.get(name)
                                )
                            else:
                                listTemp[
                                    column_index + i - 1
                                ] += " " + PEOPLE_ATTENDEES.get(name)
                        else:
                            name = attendee["email"].split(".")
                            firstName = name[0]
                            surname = name[1].split("@")[0]
                            if i < 6:
                                listTemp.insert(
                                    column_index + i, firstName + " " + surname
                                )
                            else:
                                listTemp[column_index + i - 1] += (
                                    " " + firstName + " " + surname
                                )
                        if i < 6:
                            i += 1


# Blank all the entries of the lists for a simpler and better usage of insert function.
def blanking(list, startIndex):
    while startIndex < 25:
        list.insert(startIndex, " ")
        startIndex += 1


for event in resultEventsList:

    if (
        event.get("status") == "cancelled"
        or event.get("summary") in U_EVENTS
        or "Holiday" in event.get("summary")
        or "Holidays" in event.get("summary")
    ):  # pass unwanted events
        pass
    else:
        try:
            title = event["summary"].split(
                "-"
            )  # Make sure there is no '-' (hyphen) in title besides being a separator.
            if event.get("summary") != None:
                if "Event" == title[1].strip():
                    listE = []
                    blanking(listE, 0)
                    listE.insert(0, title[0].strip())  # hub
                    if event["start"].get("dateTime") != None:
                        dateB(
                            event["start"]["dateTime"],
                            event["end"]["dateTime"],
                            listE,
                            True,
                        )
                    else:  # For all day events dateTime parameter will be replaced with date.
                        dateB(
                            event["start"]["date"], event["end"]["date"], listE, False
                        )
                    lengthT = len(event["summary"].split("-"))
                    i = 2
                    listTemp = ""
                    while i < lengthT:
                        listTemp += title[i].strip() + " "
                        i += 1
                    listE.insert(4, listTemp)  # Name of Event
                    if event.get("location") != None:
                        listE.insert(5, event["location"])  # Location
                    if event.get("attendees") != None:
                        attendeesF(
                            event["attendees"], listE, 6
                        )  # 6 is the column where the first attendee will be introduced.
                    if event.get("description") != None:
                        notes = event["description"]
                    listE.insert(12, notes)
                    listH.append(listE)
                elif "Meeting" == title[1].strip():
                    # print(title[1].strip())
                    listG = []
                    blanking(listG, 0)
                    listG.insert(0, title[0].strip())  # hub
                    if event["start"].get("dateTime") != None:
                        dateB(
                            event["start"]["dateTime"],
                            event["end"]["dateTime"],
                            listG,
                            True,
                        )
                    else:
                        dateB(
                            event["start"]["date"], event["end"]["date"], listG, False
                        )
                    lengthT = len(event["summary"].split("-"))
                    i = 2
                    listTemp = ""
                    while i < lengthT:
                        listTemp += title[i].strip() + " "
                        i += 1
                    listG.insert(4, listTemp)  # Name of Event
                    # print( event["location"])
                    if event.get("location") != None:
                        listG.insert(5, event["location"])  # Location

                    # try:
                    # print(event["attendees"])
                    if event.get("attendees") != None:
                        attendeesF(
                            event["attendees"], listG, 6
                        )  # 6 is the colun where the first attendee will be introduced.
                    # except:
                    #    print(event.get("id"))
                    # print(event["attendees"])
                    if event.get("description") != None:
                        notes = event["description"]
                    listG.insert(12, notes)
                    listJ.append(listG)
                elif "Club" == title[1].strip():
                    listT = []
                    blanking(listT, 0)
                    listT.insert(0, title[0].strip())  # hub
                    if event["start"].get("dateTime") != None:
                        dateB(
                            event["start"]["dateTime"],
                            event["end"]["dateTime"],
                            listT,
                            True,
                        )
                    else:
                        dateB(
                            event["start"]["date"], event["end"]["date"], listT, False
                        )
                    notes = ""
                    lengthT = len(event["summary"].split("-"))
                    i = 2
                    listTemp = ""
                    while i < lengthT:
                        listTemp += title[i].strip() + " "
                        i += 1
                    listT.insert(4, listTemp)  # Name of club
                    if event.get("description") != None:
                        descriptionN = event["description"]
                        descriptionN = descriptionN.split("\n")
                        i = 0
                        for desc in descriptionN:
                            # Formatted line
                            if desc == "":
                                notes += desc
                            else:
                                if ":" in desc:
                                    left = desc[: desc.index(":") + 1]
                                    if left in EVENT_INFORMATION:
                                        index = EVENT_INFORMATION.index(left)
                                        right = desc[desc.index(":") + 2 :]
                                        listT.pop(5 + index)
                                        listT.insert(5 + index, right)
                                    else:
                                        notes += desc + " "
                                else:
                                    notes += desc + " "
                    listT.insert(15, notes)
                    if event.get("attendees") != None:
                        attendeesF(
                            event["attendees"], listT, 16
                        )  # 16 is the column where the first attendee will be introduced.
                    if event.get("location") != None:
                        listT.insert(22, event["location"])  # Location
                    listTc.append(listT)
                elif "Teachers" == title[1].strip():
                    listF = []
                    blanking(listF, 0)
                    listF.insert(0, title[0].strip())  # hub
                    if event["start"].get("dateTime") != None:
                        dateB(
                            event["start"]["dateTime"],
                            event["end"]["dateTime"],
                            listF,
                            True,
                        )
                    else:
                        dateB(
                            event["start"]["date"], event["end"]["date"], listF, False
                        )
                    lengthT = len(event["summary"].split("-"))
                    i = 2
                    listTemp = ""
                    while i < lengthT:
                        listTemp += title[i].strip() + " "
                        i += 1
                    listF.insert(4, listTemp)  # Name of Event
                    if event.get("location") != None:
                        listF.insert(5, event["location"])  # Location
                    if event.get("attendees") != None:
                        attendeesF(
                            event["attendees"], listF, 6
                        )  # 6 is the column where the first attendee will be introduced.
                    notes = ""
                    if event.get("description") != None:
                        descriptionN = event["description"]
                        descriptionN = descriptionN.split("\n")
                        for desc in descriptionN:
                            # Formatted line
                            if desc == "":
                                notes += desc
                            else:
                                if ":" in desc:
                                    left = desc[: desc.index(":") + 1]
                                    if left in EVENT_INFORMATION:
                                        right = desc[desc.index(":") + 2 :]
                                        listF.insert(12, right)
                                    else:
                                        notes += desc + " "
                                else:
                                    notes += desc + " "
                    listF.insert(13, notes)
                    listI.append(listF)
                elif "Workshop" == title[1].strip():
                    listA = []
                    blanking(listA, 0)
                    listA.insert(0, title[0].strip())  # hub
                    if event["start"].get("dateTime") != None:
                        dateB(
                            event["start"]["dateTime"],
                            event["end"]["dateTime"],
                            listA,
                            True,
                        )
                    else:
                        dateB(
                            event["start"]["date"], event["end"]["date"], listA, False
                        )
                    listA.insert(5, title[2].strip())  # school
                    listA.insert(6, title[3].strip())  # ks
                    lengthT = len(event["summary"].split("-"))
                    i = 4
                    listTemp = ""
                    while i < lengthT:
                        listTemp += title[i].strip() + " "
                        i += 1
                    listA.insert(7, listTemp)  # workshop
                    notes = ""
                    if event.get("description") != None:
                        descriptionN = event["description"].replace("<br>", "\n")
                        descriptionN = descriptionN.split("\n")
                        i = 0
                        for desc in descriptionN:
                            # Formatted line
                            if desc == "":
                                notes += desc
                            else:
                                if ":" in desc:
                                    left = desc[: desc.index(":") + 1]
                                    if left in EVENT_INFORMATION:
                                        index = EVENT_INFORMATION.index(left)
                                        right = desc[desc.index(":") + 2 :]
                                        if index != 11:
                                            listA.pop(
                                                8 + index
                                            )  # Entries can't be converted to int because of "3/4" examples.
                                            listA.insert(8 + index, right)
                                        else:  # School Number
                                            listA.pop(4)
                                            listA.insert(4, right)
                                    else:
                                        notes += desc + " "
                                else:
                                    notes += desc + " "
                    listA.insert(18, notes)
                    if event.get("attendees") != None:
                        attendeesF(
                            event["attendees"], listA, 19
                        )  # 19 is the column where the first attendee will be introduced.
                        # This will add to Officers' Workshop sheet .
                        workshops(event["attendees"], listA)
                    listC.append(listA)
                elif title[1].strip() in SPECIAL_EVENTS:
                    listS = []
                    blanking(listS, 0)
                    listS.insert(0, title[0].strip())  # hub
                    if event["start"].get("dateTime") != None:
                        dateB(
                            event["start"]["dateTime"],
                            event["end"]["dateTime"],
                            listS,
                            True,
                        )
                    else:  # For all day events dateTime parameter will be replaced with date.
                        dateB(
                            event["start"]["date"], event["end"]["date"], listS, False
                        )
                    lengthT = len(event["summary"].split("-"))
                    i = 1
                    listTemp = ""
                    while i < lengthT:
                        listTemp += title[i].strip() + " "
                        i += 1
                    listS.insert(4, listTemp)  # Name of Event
                    if event.get("location") != None:
                        listS.insert(5, event["location"])  # Location
                    if event.get("attendees") != None:
                        attendeesF(
                            event["attendees"], listS, 6
                        )  # 6 is the column where the first attendee will be introduced.
                    if event.get("description") != None:
                        notes = event["description"]
                    listS.insert(12, notes)
                    listSS.append(listS)
                else:  # incorect title
                    listB = []
                    if event["start"].get("dateTime") != None:
                        dateB(
                            event["start"]["dateTime"],
                            event["end"]["dateTime"],
                            listB,
                            True,
                        )
                    else:
                        dateB(
                            event["start"]["date"], event["end"]["date"], listB, False
                        )

                    listB.insert(3, event["summary"])  # Event Title
                    blanking(listB, 4)
                    if event.get("attendees") != None:
                        attendeesF(
                            event["attendees"], listB, 4
                        )  # 4 is the column where the first attendee will be introduced.
                    if event.get("description") != None:
                        listB.insert(10, event["description"])  # Notes
                    listD.append(listB)
        except:  # Events with errors: Usually there format of the title is wrong, making the event to fall into wrong category(sheet).
            # print(event)
            listB = []
            if event["start"].get("dateTime") != None:
                dateB(event["start"]["dateTime"], event["end"]["dateTime"], listB, True)
            else:
                dateB(event["start"]["date"], event["end"]["date"], listB, False)

            listB.insert(3, event["summary"])  # Event Title
            blanking(listB, 4)
            if event.get("attendees") != None:
                attendeesF(
                    event["attendees"], listB, 4
                )  # 4 is the column where the first attendee will be introduced.
            if event.get("description") != None:
                listB.insert(10, event["description"])  # Notes
            listD.append(listB)


# Sort the lists by Hub.
# Put the order of the Hubs. For alphabetical order replace the value of the key with itemgetter(0)
SORT_ORDER = {
    "HUB1": 0,
    "HUB2": 1,
    "HUB3": 2,
    "HUB4": 3,
    "HUB5": 4,
    "HUB6": 5,
    "HUB7": 6,
    "HUB8": 7,
    "HUB9": 8,
    "Hub?": 9,
    "HUB1 / HUB10": 10,
}
# An error should not occur unless the somebody added a hub in HUBS but didn't added it in SORT_ORDER
try:
    listC = sorted(listC, key=lambda val: SORT_ORDER[val[0]])
    listH = sorted(listH, key=lambda val: SORT_ORDER[val[0]])
    listI = sorted(listI, key=lambda val: SORT_ORDER[val[0]])
    listJ = sorted(listJ, key=lambda val: SORT_ORDER[val[0]])
    listTc = sorted(listTc, key=lambda val: SORT_ORDER[val[0]])
    listSS = sorted(listSS, key=lambda val: SORT_ORDER[val[0]])
except:
    pass
listMi = copy.deepcopy(listC)

# Write the lists in Excel.
def writeE(list, worksheet):
    row = 1
    for event in list:
        column = 0
        for entry in event:
            if isinstance(entry, datetime.datetime):
                worksheet.write_datetime(row, column, event[column], cell_format)
            else:
                worksheet.write(row, column, event[column], alignment)
            column += 1
        row += 1


# Switch columns for Booking sheet.
for event in listMi:
    j = 9  # Position for attendees to be inserted
    i = 0
    listT = []  # List of attendees
    t = 0  # Attendee index
    while t < 6:
        listT.append(event[19])
        event.pop(19)
        t += 1
    while j <= 14:
        event.insert(j, listT[i])
        i += 1
        j += 1

listOfLists = [
    [listC, worksheet],
    [listD, worksheetM],
    [listMi, worksheetB],
    [listH, worksheetE],
    [listI, worksheetT],
    [listJ, worksheetMe],
    [listTc, worksheetTc],
    [listSS, worksheetS],
]
for list in listOfLists:
    writeE(list[0], list[1])

workbook.close()
