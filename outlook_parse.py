import sqlite3
import os
import win32com
import win32com.client
from datetime import date, datetime, timedelta
from time import sleep, time

today = date.today()
yesterday = (today - timedelta(days = 1)).strftime("%d-%m-%y")
today = today.strftime("%d-%m-%y")
outlook = win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")


#Database Setup and config
db = sqlite3.connect("emails.db")
db.execute("""CREATE TABLE IF NOT EXISTS errors (ID INTEGER PRIMARY KEY, ConversionFailed integer, UnsupportedAttachments integer,
        BatchValidationFailed integer,NotRecieved integer, AttachmentNoExtension integer,InvalidFilename integer, Date text)""")


print("\nCalculating please wait ...")


inbox = outlook.GetDefaultFolder(6).Folders("test")
error_folder = outlook.GetDefaultFolder(6).Folders("2021")
messages = inbox.items
recieved_dt = datetime.now() - timedelta(days=1)
recieved_dt = recieved_dt.strftime('%d/%m/%Y')
todays_date = datetime.now().strftime('%d/%m/%Y')

previous_date = (datetime.now() - timedelta(days=3)).strftime('%d/%m/%Y')

errors = ("Unsupported Attchments", "Invalid Filename", "Conversion Failed", "Attachment with no extension")
print("This is the previous date", previous_date)
messages = messages.Restrict("[Subject] = 'Conversion Failed'")

todays_date = datetime.now().strftime('%m/%d/%Y %H:%M %p')
previous_date = (datetime.now() - timedelta(days=3)).strftime('%m/%d/%Y %H:%M %p')
run = False

count = 0
#Checks emails from 3 days ago - good for Monday getting emails from the weekend
def main():
    global count
    for i in messages:
        if i.subject in errors:
            print(i.subject, i.ReceivedTime.strftime('%d/%m/%Y'))
            i.Move(error_folder)
            run = True
            count +=1
        
        else:
            run = True
            continue

def hourly_check():
    for i in messages:
        pass


cur_time = todays_date.split()
cur_time = " ".join(cur_time[1:])

def print_toscreen():
    if count == 0:
        print(f"Check complete at {cur_time}!\n\nNo new errors were found")
    else:
        print(f"\n\n\nCheck Complete at {cur_time}!\n\nAll EMG Errors have moved to the 2021 folder ...")

print_toscreen()

def checkover():
    pass


if run == True:
    print("THIS IS THE NEXT PART OF THE PROGRAM")



if __name__ == "__main__":
    pass