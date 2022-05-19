from xlwt import Workbook
from time import sleep
from os import system

# Intro Animation
for _ in range(0, 2):
    system("cls")
    print("//// Mom's Program ////")
    sleep(.2)
    system("cls")
    print("~~~~ Mom's Program ~~~~")
    sleep(.2)
    system("cls")
    print("\\\\\\\\ Mom's Program \\\\\\\\")
    sleep(.2)
    system("cls")
    print("|||| Mom's Program ||||")
    sleep(.2)

# Try to import file
found = False
while not found:
    filename = input("Type the exact file name: ")
    try:
        with open(f"Input/{filename}", "r") as f:
            email = f.read()
        found = True
    except FileNotFoundError:
        print("File not found...")

# Variables
newFilename = filename.replace(".txt", "")
newFilename = newFilename.replace('Shoutbomb', '')

queries = {
    "Hold notices sent for the month": 0,
    "Hold cancel notices sent for the month": 0,
    "Overdue notices sent for the month": 0,
    "Overdue items eligible for renewal, notices sent for the month": 0,
    "Overdue items ineligible for renewal, notices sent for the month": 0,
    "Overdue items renewed successfully by patrons for the month": 0,
    "Overdue items unsuccessfully renewed by patrons for the month": 0,
    "Renewal notices sent for the month": 0,
    "Items eligible for renewal notices sent for the month": 0,
    "Items ineligible for renewal notices sent for the month": 0,
    "Items renewed successfully by patrons for the month": 0,
    "Items unsuccessfully renewed by patrons for the month": 0,
    }

libraries = {
    "Atkinson": 0, "Bay View": 0, "Villard": 0, "Wash Park": 0, "Capitol": 0,
    "Mitchell St.": 0, "Zablocki": 0, "Center St.": 0,
    "Hales Corners": 0, "Whitefish Bay": 0, "Shorewood": 0, "Cudahy": 0,
    "North Shore": 0, "Brown Deer": 0, "Tippecanoe": 0, "St. Francis": 0,
    "Good Hope": 0, "West Allis": 0, "Wauwatosa": 0, "Oak Creek": 0,
    "West Milwaukee": 0, "King": 0, "Greendale": 0, "Greenfield": 0,
    "East": 0, "South Milwaukee": 0, "Franklin": 0, "Central": 0,
    }

workbook = Workbook()
splittedEmail = email.split("=TOTALS BY BRANCH=")[0]


def parse(data, query):
    for line in data.splitlines():
        for key in query.keys():
            if key in line:
                newLine = line.split(" = ")
                newLine = int(newLine[1])
                query[key] = newLine
    return query


# First Sheet
totalsByBranch = workbook.add_sheet(f"Totals {newFilename}")

emailText = splittedEmail.split("=TOTALS=")[0]
queriesList = list(queries.keys())
for query in queriesList:
    totalsByBranch.write(int(queriesList.index(query)+1), 0, query)
row = 0
column = 0
for branch in emailText.split("Branch:: "):
    for library in libraries:
        row = 0
        if library in branch:
            column += 1
            totalsByBranch.write(0, column, library)
            libQueries = parse(branch, queries.copy())
            for query in libQueries.values():
                row += 1
                totalsByBranch.write(row, column, query)


row = 0 
column += 1
totals = parse(splittedEmail.split("=TOTALS=")[1], queries.copy())
totalsByBranch.write(row, column, f"Totals")
for query in totals.values():
    row += 1
    totalsByBranch.write(row, column, query)


# Second Part
column = 1
textNotices = workbook.add_sheet(f"Text Notices Sent {newFilename}")
textNotices.write(1, 0, "Total Text Notices")
splittedEmail = email.split("=TOTALS BY BRANCH=")[1]
emailText = splittedEmail.split("=TOTALS OF REGISTERED PATRON BY BRANCH=")[0]
values = parse(emailText, libraries)
for line in emailText.splitlines():
    for library in libraries:
        if library in line:
            textNotices.write(0, column, library)
            textNotices.write(1, column, values[library])
            column += 1

# Third Part
column = 1
registeredUsers = workbook.add_sheet(f"Registered Patrons {newFilename}")
registeredUsers.write(1, 0, "Total Registered Users")
emailText = splittedEmail.split("=TOTALS OF REGISTERED PATRON BY BRANCH=")[1]

# setup custom parsing
libraryCopy = libraries.copy()
for line in emailText.splitlines():
        for key in libraries.keys():
            if key in line:
                newLine = line.split(" has ")[1]
                newLine = newLine.replace(" registered patrons for text notices", "")
                newLine = int(newLine)
                libraryCopy[key] = newLine
                
for line in emailText.splitlines():
    for library in libraries:
        if library in line:
            registeredUsers.write(0, column, library)
            registeredUsers.write(1, column, libraryCopy[library])
            column += 1
            

# Save workbook
workbook.save(f"Output/{filename.replace('.txt', '.xls')}")
print("Saved Successfully...")
