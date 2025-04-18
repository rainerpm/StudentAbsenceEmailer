####################################################################################
# 4/14/2025  Version 1.1
####################################################################################
import csv
import win32com.client  # pip install pywin32 (close and reopen Python after install) [for email using Outlook Windows app (https://github.com/mhammond/pywin32)] (Thonny install pywin32 package)
import re
import random
from datetime import datetime  # module is in python standard library
from datetime import date  # module is in python standard library
from jinja2 import Template  # pip install jinja2
import pkgutil   # library that allows exe file to include and read the csv file data
import io        # used with above import
import sys
import os
from   pathlib  import Path      # module is in python standard library
import subprocess

version = "1.1c"
versionDate = "4/18/25"

## To generate an EXE file for Windows (in a cmd tool) using pyinstaller
## (1) cd C:\Users\E151509\My Drive\My LASA\misc\tools\studentAbsenceEmailer
## (2) pyinstaller --onefile --add-data "studentAbsenceEmailerData2024-25.csv;." studentAbsenceEmailer.py
## (3) move studentAbsenceEmailer.exe from 'dist' folder to 'Latest EXE Release linked in documentation Google Doc' folder

## To generate a MAC OS X App Bundle using pyinstaller (might have to use 'pip3 install pyinstaller)
## (1) comment out import win32.com.client line above
## (2) pyinstaller --windowed --onefile --osx-bundle-identifier "<org.austinisd.studentAbsenceEmailer>" --add-data "studentAbsenceEmailerData2024-25.csv:." studentAbsenceEmailer.py
## (3) move studentAbsenceEmailer.exe from 'dist' folder to 'Latest Apple Release linked in documentation Google Doc' folder

csvStudentsFile = "studentAbsenceEmailerData2024-25.csv"

defaultEmailFooter = """
This email was sent by the Student Absence Emailer (SAE) program. Programs have bugs (especially ones written by a CS teacher).  Please let rainer.mueller@austinisd.org know if something looks awry. Documentation @ https://tinyurl.com/LASAStudentAbsenceEmailer
"""

ignorePeriods = ["A-AFA","A-AFB","A-BFA","B-AFB","B-BFB"]  # ignore students with one of these periods
ignorePeriodTypes = ["OFF PERIOD","OFFICE AIDE"]  # ignore students with one of these period types

ADays = ["Monday", "Wednesday"]
BDays = ["Tuesday", "Thursday"]
XDays = ["Friday"]
periodsADay = "1,2,3,4"
periodsBDay = "5,6,7,8"
periodsAdvADay = "1,Adv,2,3,4"
periodsAdvBDay = "5,Adv,6,7,8"


class bcolors:
    HEADER,BLUE,CYAN,GREEN,WARNING,RED,ENDC,BOLD,UNDERLINE,LIGHTGRAY,ORANGE,BLACK,BGGREEN,BGRED,BGYELLOW,BGCYAN = '\033[95m','\033[94m','\033[96m','\033[92m','\033[93m','\033[91m','\033[0m','\033[1m','\033[4m','\033[37m','\033[33m','\033[30m','\033[42m','\033[41m','\033[43m','\033[46m'

def emailWithOutlookPC(email_recipient, email_subject, email_message):
    outlook = win32com.client.Dispatch("Outlook.Application")
    email = outlook.CreateItem(0)
    email.To = email_recipient
    email.Subject = email_subject
    email.HTMLBody = email_message
    email.send  # SEND THE EMAIL
    return True

# does NOT work with HTML (only with text emails)
def emailWithOutlookApple(recipient, subject, body):
    apple_script = f'''
    tell application "Microsoft Outlook"
        set newMessage to make new outgoing message with properties {{subject:"{subject}"}}
        set content of newMessage to "{body}"
        make new recipient at newMessage with properties {{email address:{{address:"{recipient}"}}}}
        send newMessage
    end tell
    '''
    subprocess.run(['osascript', '-e', apple_script])

def getPeriod(periodStr):
    if periodStr.startswith("A-01"):
        return "1"
    elif periodStr.startswith("A-02"):
        return "2"
    elif periodStr.startswith("A-03"):
        return "3"
    elif periodStr.startswith("A-04"):
        return "4"
    elif periodStr.startswith("B-05"):
        return "5"
    elif periodStr.startswith("B-06"):
        return "6"
    elif periodStr.startswith("B-07"):
        return "7"
    elif periodStr.startswith("B-08"):
        return "8"
    elif periodStr.startswith("A-ADV"):
        return "Adv"
    else:
        return "???"

# suggested by chatgpt.com to solve a problem with the exe file reading in the csv file
def get_data_file_path(filename):
    if getattr(sys, 'frozen', False):
        # If running from PyInstaller bundle
        base_path = sys._MEIPASS
    else:
        # Running in normal Python
        base_path = os.path.abspath(".")

    return os.path.join(base_path, filename)

# chatgpt.com prompt "Python code that checks that a datetime object is between the prior August and the upcoming June "
def is_between_prior_aug_and_upcoming_june(date_to_check: datetime) -> bool:
    today = datetime.today()

    if today.month >= 8:  # August to December
        august_start = datetime(today.year, 8, 1)
        june_end = datetime(today.year + 1, 6, 30, 23, 59, 59)
    else:  # January to June 8th
        august_start = datetime(today.year - 1, 8, 1)
        june_end = datetime(today.year, 6, 8, 23, 59, 59)

    return august_start <= date_to_check <= june_end



def main():
    print(f'Version {version} ({versionDate}) for the 2024-25 school year.')
    print(f'Documentation at https://tinyurl.com/LASAStudentAbsenceEmailer')
    #print(f'Reading in {csvStudentsFile} file')

    windows = False
    apple = False
    if sys.platform == "darwin":
        #print("Running on macOS (Mac)")
        apple = True
    elif sys.platform == "win32":
        #print("Running on Windows (PC)")
        windows = True
    else:
        print(f'{bcolors.RED}Warning!!!{bcolors.ENDC} Can not determine if running on a Windows PC or a Mac.')               
        
    
    #################################################################
    ### read data from spreadsheet's csv file
    #################################################################
    students = {}
    rowCount = 0
    
    cvsPath = get_data_file_path(csvStudentsFile)
    cvsFileDateTime = datetime.fromtimestamp(Path(cvsPath).stat().st_mtime)
    if not is_between_prior_aug_and_upcoming_june(cvsFileDateTime):
        print(f'{bcolors.RED}Warning!!!{bcolors.ENDC} {csvStudentsFile} from {cvsFileDateTime.strftime("%b %d, %Y")} is not for this school year.')               
    with open(cvsPath, newline='', encoding='utf-8') as csvfile:
        csvReader = csv.reader(csvfile)       
        for row in csvReader:
            rowCount += 1
            if rowCount >= 2:
                studentID = row[0].strip()
                studentName = row[1].strip()
                periodType = row[2].strip()
                period = row[3].strip()
                teacher = row[4].strip()
                email = row[5].strip()

                continue1 = False
                for p in ignorePeriods:
                    if period.startswith(p):
                        continue1 = True
                if continue1:
                    continue
                period = getPeriod(period)
                if period == "???":
                    print(row)
                    print(f"  {bcolors.RED}Warning!!!{bcolors.ENDC} row {rowCount} Student ID {studentID} has unrecognized period {period}. Please update getPeriod() function.")
                continue2 = False
                for pt in ignorePeriodTypes:
                    if periodType.startswith(pt):
                        continue2 = True
                if continue2:
                    continue
                if not email:
                    print(f"  {bcolors.RED}Warning!!!{bcolors.ENDC} row {rowCount} Student ID {studentID} for {teacher} does NOT have a teacher email. Skipping student.")
                    continue
                if studentID in students:
                    if studentName != students[studentID][0].split("_")[0]:
                        print(f'  {bcolors.RED}Warning!!!{bcolors.ENDC} row {rowCount} Student ID {studentID} for name {studentName} already exists with name {students[studentID][0].split("_")[0]}.')
                    emailDic = students[studentID][1]
                    if period in emailDic:
                        print(f"  {bcolors.RED}Warning!!!{bcolors.ENDC} row {rowCount} Student ID {studentID} already has a previous period {period}.")
                    else:
                        emailDic[period] = email
                        students[studentID] = (studentName + "_" + studentID, emailDic)
                else:
                    students[studentID] = (studentName + "_" + studentID,{period: email})
    # pprint(students)

    #################################################################
    ### Subject
    #################################################################
    emailSubject = input(f"\n{bcolors.BOLD}Enter email subject:{bcolors.ENDC} ").strip()
    emailSubject = '[SAE] ' + emailSubject

    #################################################################
    ### Dates, Periods, and optional time(s)
    #################################################################
    print(f"\n{bcolors.BOLD}Enter mm/dd/yy #,#,#,# below (# = 1-8 or Adv).{bcolors.ENDC}")
    print("Press <ENTER> on an empty line to finish.")
    lines = []
    todaysDate = date.today()
    while True:
        line = input().strip()
        matchResponse = re.match(r"^(0?[1-9]|1[0-2])/(0?[1-9]|[12][0-9]|3[01])/\d{2}(?: ([1-8]|Adv|adv)((leaving|returning)@(0?[1-9]|1[0-2]):[0-5][0-9](am|pm|AM|PM))?(,([1-8]|Adv|adv)((leaving|returning)@(0?[1-9]|1[0-2]):[0-5][0-9](am|pm|AM|PM))?)*)?$", line)
        if line:
            if matchResponse:
                lines.append(line)
            else:
                print(f"  {bcolors.RED}Warning!!!{bcolors.ENDC} Disregarding invalid input {bcolors.RED}{line}{bcolors.ENDC}")
        else:
            break
    dates = []
    classDatePeriods = []
    periodsMissedCount = 0
    for datePeriod in lines:
        if " " in datePeriod:
            dateStr, periodAndTime = datePeriod.split()
        else:
            dateStr = datePeriod
            periodAndTime = None
        dateTimeObject = datetime.strptime(dateStr, "%m/%d/%y")  # Convert to datetime object
        dayOfTheWeek = dateTimeObject.strftime("%A")
        dateObject = dateTimeObject.date()
        daysFromNow = (dateObject - todaysDate).days
        if not periodAndTime:  # periodAndTime == None
            if dayOfTheWeek in XDays:
                while True:
                    response = input(f"  Is {dayOfTheWeek} {dateStr} an A-day or a B-day (answer 'a' or 'b')? ").strip().lower()
                    if response == "a":
                        periodAndTime = periodsAdvADay
                        break
                    elif response == "b":
                        periodAndTime = periodsAdvBDay
                        break
                    else:
                        print("  Please respond with either 'a' or 'b'.")
            elif dayOfTheWeek in ADays:
                periodAndTime = periodsADay
            elif dayOfTheWeek in BDays:
                periodAndTime = periodsBDay
        pts = periodAndTime.split(",")
        # handle only 1 period being entered
        if len(pts) == 1:
            onlyPeriod = pts[0].split('@')[0].strip("leaving")
            if dayOfTheWeek in XDays:
                while True:
                    response = input(f"  Is {dayOfTheWeek} {dateStr} an A-day or a B-day (answer 'a' or 'b')? ").strip().lower()
                    if response == "a":
                        periodAndTime = periodsAdvADay[periodsAdvADay.find(onlyPeriod):]
                        break
                    elif response == "b":
                        periodAndTime = periodsAdvBDay[periodsAdvBDay.find(onlyPeriod):]
                        break
                    else:
                        print("  Please respond with either 'a' or 'b'.")
            elif dayOfTheWeek in ADays:
                periodAndTime = periodsADay[periodsADay.find(onlyPeriod):]
            elif dayOfTheWeek in BDays:
                periodAndTime = periodsBDay[periodsBDay.find(onlyPeriod):]
            periodAndTime = periodAndTime.replace(onlyPeriod,pts[0])   # put the leaving@ or returning@ back in
        pts = periodAndTime.split(",")
        periodsMissedCount += len(pts)
        periodsStr = ""
        printPeriodsStr = ""
        datePeriods = []
        for pt in pts:
            leaveReturnStr = None
            timeStr = None
            if pt.find("@") != -1:
                periodLeaveReturnStr, timeStr = pt.split("@")
                matchResponse = re.match(r"^(\w+)(leaving|returning)$", periodLeaveReturnStr)
                periodStr = matchResponse.group(1)
                leaveReturnStr = matchResponse.group(2)
            else:
                periodStr = pt                     
            datePeriods.append((periodStr, leaveReturnStr, timeStr))
            if leaveReturnStr:
                printPeriodsStr += f"{periodStr}({leaveReturnStr} at {timeStr}) "
            else:
                printPeriodsStr += f"{periodStr} "
        print(
            f'  in {daysFromNow} day{"s" if daysFromNow > 1 else ""}, on {dayOfTheWeek} {dateStr}, periods: {printPeriodsStr}'
        )
        dates.append(dateStr)
        classDatePeriods.append(datePeriods)

    #################################################################
    ### Student IDs
    #################################################################
    print(f"{bcolors.BOLD}\nEnter student IDs below (one per line).{bcolors.ENDC}")
    print("Press <ENTER> on an empty line to finish.")
    lines = []
    while True:
        line = input().strip()
        if line:
            if line in lines:
                print(f"  Duplicate ID {line} found. Disregarded.")
            else:
                lines.append(line)
        else:
            break
    studentIDs = lines
    studentCount = len(lines)

    #################################################################
    ### Store all the data in the emails dictionary
    #################################################################
    emailAbsencesHTML = ''
    emails = {}  
    studentsNotFound = {}
    #print(f'{dates = }')
    for i in range(len(dates)):
        classPeriods = classDatePeriods[i]
        for j in range(len(classPeriods)):
            for studentID in studentIDs:
                if studentID in students:
                    studentName = students[studentID][0]
                    classPeriod = classPeriods[j][0]
                    if classPeriods[j][0] in students[studentID][1]:
                        teacherEmailForPeriod = students[studentID][1][classPeriods[j][0]]
                    else:
                        continue
                    if teacherEmailForPeriod in emails:
                        if dates[i] in emails[teacherEmailForPeriod]:
                            if classPeriods[j][0] in emails[teacherEmailForPeriod][dates[i]]:
                                emails[teacherEmailForPeriod][dates[i]][classPeriods[j][0]].append(studentName)
                            else:
                                emails[teacherEmailForPeriod][dates[i]][classPeriods[j][0]] = [(classPeriods[j][1],classPeriods[j][2]),studentName]
                        else:
                            emails[teacherEmailForPeriod][dates[i]] = {classPeriods[j][0] : [(classPeriods[j][1],classPeriods[j][2]),studentName]}
                    else:
                        emails[teacherEmailForPeriod] = {dates[i] : {classPeriods[j][0] : [(classPeriods[j][1],classPeriods[j][2]),studentName]} }
                else:
                    studentsNotFound[studentID] = True
    for key in studentsNotFound:
        print(f'{bcolors.RED}Warning!!!{bcolors.ENDC} Student ID {key} was not found in {csvStudentsFile}!!!')

    #################################################################
    ### transfer data into the list emailsList
    #################################################################    
    emailsList = []
    for email in emails:
        emailList = [email]
        dateList = []
        for dateStr in emails[email]:
            periodList = []
            for period in emails[email][dateStr]:
                periodList.append((period, emails[email][dateStr][period]))
            periodList.sort()
            dateList.append([dateStr, periodList])
        dateList.sort()
        emailList.append(dateList)
        emailsList.append(emailList)
    emailsList.sort()
    #pprint(emailsList)

    #################################################################
    ### Email Message
    #################################################################
    print(f"{bcolors.BOLD}Enter email message below.{bcolors.ENDC} ")
    print("Press <ENTER> on an empty line to finish.")
    lines = []
    while True:
        line = input().strip()
        if line:
            lines.append(line)
        else:
            break
    emailMessage = "\n".join(lines)

    #################################################################
    ### On behalf
    #################################################################
    onBehalfOfStatement = ''
    onBehalfOfStatementText = ''
    onBehalfOfName = input(f"{bcolors.BOLD}Enter name on whose behalf the emails are being sent (or <Enter> to skip):{bcolors.ENDC} ").strip()
    if onBehalfOfName:
        onBehalfOfStatement = f'<p>This email was sent by SAE on behalf of {onBehalfOfName}.</p>'
        onBehalfOfStatementText = f'This email was sent by SAE on behalf of {onBehalfOfName}.\n'

    #################################################################
    ### Send the emails
    #################################################################    
    emailHTMLTemplate = Template(
        """
    <html>
    <head>
        <title>{{ title }}</title>
    </head>
    <body>
    {{onBehalfOfStatement}}
    <p>{{emailMessage}}</p>
    <p>The students listed below will miss <b>your</b> classes.</p>
    {{emailAbsencesHTML}}
    <p>{{defaultEmailFooter}}</p>
    </body>
    </html>
    """
    )
    
    emailTextTemplate = Template(
        """
{{onBehalfOfStatementText}}
{{emailMessage}}
    
The students listed below will miss YOUR classes.
    
{{emailAbsencesText}}
    
{{defaultEmailFooter}}
    """
    )

    testEmailRecipient = input(f"{bcolors.BOLD}Test email address (or <Enter> to skip):{bcolors.ENDC} ").strip()
    
    emailCount = 0
    studentAbscenceCount = 0
    sendTest = True
    sendingEmails = False
    for emailList in emailsList:
        #pprint(emailList)
        emailCount += 1
        emailAddress = emailList[0]
        periodStudentStr = ""
        emailAbsencesHTML = ""
        emailAbsencesText = ""
        for dateList in emailList[1]:
            dateStr = dateList[0]
            dateTimeObject = datetime.strptime(
                dateStr, "%m/%d/%y"
            )  # Convert to datetime object
            dateObject = dateTimeObject.date()
            daysFromNow = (dateObject - todaysDate).days
            dayOfTheWeek = dateTimeObject.strftime("%A")
            periodsList = dateList[1]
            #emailAbsencesHTML += f'<p><strong><h3 style="display:inline;">{dayOfTheWeek} {dateStr}</h3></strong>&nbsp;in {daysFromNow} day{"s" if daysFromNow != 1 else ""}</p>'
            emailAbsencesHTML += f'<h3 style="display:inline;">{dayOfTheWeek} {dateStr}</h3>'
            emailAbsencesText += f'\n{dayOfTheWeek} {dateStr}\n'
            for periodTuple in periodsList:
                period = periodTuple[0]
                timeTuple = periodTuple[1][0]
                studentList = periodTuple[1][1:]
                emailAbsencesHTML += "<ul>"
                emailAbsencesHTML += f"<h4>Period {period}"
                emailAbsencesText += f"\nPeriod {period}"
                if timeTuple[0]:
                    emailAbsencesHTML += f" ({timeTuple[0]} at {timeTuple[1]})"
                    emailAbsencesText += f" ({timeTuple[0]} at {timeTuple[1]})"
                emailAbsencesHTML += "</h4>"
                emailAbsencesText += "\n"
                periodStudentStr += f"Period {period}({len(studentList)}) "
                studentAbscenceCount += len(studentList)
                studentList.sort()
                # Table of students
                emailAbsencesHTML += "<table>"
                for student in studentList:
                    studentName, studentID = student.split("_")
                    emailAbsencesHTML += f'<tr><td style="padding-right: 15px;">{studentName}</td> <td>{studentID}</td></tr>'
                    emailAbsencesText += f'\t{studentName[0:27]:27} {studentID}\n'
                emailAbsencesHTML += "</table>"

                emailAbsencesHTML += "</ul>"
        # render the HTML email
        emailBodyHTML = emailHTMLTemplate.render(
            title="SAE email",
            onBehalfOfStatement=onBehalfOfStatement,
            emailMessage=emailMessage,
            emailAbsencesHTML=emailAbsencesHTML,
            defaultEmailFooter=defaultEmailFooter,
        )
        emailBodyText = emailTextTemplate.render(
            onBehalfOfStatementText=onBehalfOfStatementText,
            emailMessage=emailMessage,
            emailAbsencesText=emailAbsencesText,
            defaultEmailFooter=defaultEmailFooter,
        )

        if sendTest and testEmailRecipient:
            sendTest = False
            print(f"Sending a test email to {testEmailRecipient}. Check to make sure email looks OK.")
            if windows:               
                emailWithOutlookPC(testEmailRecipient, emailSubject, emailBodyHTML)
            elif apple:
                emailWithOutlookApple(testEmailRecipient, emailSubject, emailBodyText)
            
        while not sendingEmails:
            response = input(f"{bcolors.BOLD}Send student absence email to teachers (answer 'y' or 'n')?{bcolors.ENDC} ").strip().lower()
            if response == 'n':
                print("  Exiting program!!!")
                sys.exit()
            elif response == 'y':
                sendingEmails = True
                break
            else:
                print("  Please respond with either 'y' or 'n'.")

        #### SEND THE EMAILS
        print(f"  Sending email to {emailAddress}  {periodStudentStr}")
        if windows:
            emailWithOutlookPC(emailAddress,emailSubject,emailBodyHTML)
        elif apple:
            emailWithOutlookApple(emailAddress,emailSubject,emailBodyText)

    print(f"\nDONE!!! Sent {emailCount} emails ({studentCount} students missed {studentAbscenceCount} student-periods. {studentAbscenceCount/periodsMissedCount:.2f} students/period.)")

    input("Press <Enter> to close window")

if __name__ == '__main__':
    main()