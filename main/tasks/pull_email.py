import os
import datetime
import time
from .utils.wrapper import *
import win32com.client # pip install -U pypiwin32
        
def read_message(absolute_file):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    message = outlook.OpenSharedItem(absolute_file)
    return message

def extract_email_tuple():
    for file in os.scandir(input_fldr):
        f = file.name

        log("Found file: " + f)
        msg = read_message(input_fldr + "\\" + file.name)

        mail_sent_on = msg.SentOn.strftime("%Y/%m/%d/%H/%M/%S")
        mail_subject = msg.Subject.strip()
        mail_content = []

        for line in msg.Body.strip().split("\n"):
            line = line.strip()
            if (line):
                mail_content.append(line)

        log("!----- Sent On: \"" +  mail_sent_on + "\"")
        log("!----- Subject: \"" +  mail_subject + "\"")
        log("!----- Body: \"" +     str(mail_content) + "\"")

        return (mail_sent_on, mail_subject, mail_content)

def populate_pointers(email_tuple):
# Vendor Name
    for vendor in vendors:
        for word in email_tuple[1].lower().split(" "):
            passed = word.lower().strip() == vendor.lower().strip()

            if passed:
                log("Found Vendor Name in Mail Subject: " + vendor)
                print("Vendor Name: " + vendor)
                local_pointers["VNDR_NAME"] = vendor
    
    if local_pointers["VNDR_NAME"] == "": 
        local_pointers["VNDR_NAME"] = input("Enter the vendor name:  ")

# Maintenance ID
    for word in email_tuple[1].split():
        if word.startswith("CHG") or word.startswith("CRQ") or word.startswith("#") or word.startswith("CM"):
            log("Found Maintenance ID in Mail Subject: " + word)
            print("Maintenance ID: " + word)
            local_pointers["MAINT_ID"] = str(word)
        else:
            potential = [int(s) for s in word.split() if s.isdigit()]
            if potential == None or potential == "": continue
            if len(potential) <= 3 or potential is datetime.datetime.today().year: continue

            use_potential = input("Potential Maintenance ID: " + potential + ", use this? Yes or No")

            if use_potential.lower() == "yes": local_pointers["MAINT_ID"] = str(potential)
            else: local_pointers["MAINT_ID"] = input("Enter the maintenance ID:  ")

    if local_pointers["MAINT_ID"] == "": local_pointers["MAINT_ID"] = input("Enter the maintenance ID:  ")

# Maintenance Description
# Find a way to effectively assume the description.
    # action_words = ["summary", "description", ""]
    # for line in email_tuple[2]:

    local_pointers["MAINT_DESC"] = input("Enter the maintenance description:  ")

# Maintenance Type
    for type in type_four:
        if type in local_pointers["MAINT_DESC"]:
            use_potential = input("Potential Maintenance Type: " + type + ", use this? Yes or No:  ")

            if use_potential.lower() == "yes": local_pointers["MAINT_TYPE"] = type
            else: local_pointers["MAINT_ID"] = input("Enter the maintenance type:\n\tOptions: " + str(type_four).replace("'", "").replace("[", "").replace("]", "").replace("\\\\", "\\") + ".\n\t ::  ")

    if local_pointers["MAINT_TYPE"] == "": local_pointers["MAINT_TYPE"] = input("Enter the maintenance type:\n\tOptions: " + str(type_four).replace("'", "").replace("[", "").replace("]", "").replace("\\\\", "\\") + ".\n\t ::  ")

# Maintenance Received
    dt = email_tuple[0].split("/")
    print("Date Recieved has been located.")
    local_pointers["MAINT_DATE_REC"] = "-".join(dt[0:3])
    local_pointers["MAINT_HOUR_REC"] = dt[3]
    local_pointers["MAINT_MIN_REC"] = dt[4]

# Outage Duration
    
    outage = input("Enter the outage duration (6 hours, 30 minutes, etc):  ").split(" ")
    local_pointers["MAINT_OUTAGE_NUM"] = outage[0].strip()
    local_pointers["MAINT_OUTAGE_UNT"] = outage[1].strip()

# Work Location
    location = input("Enter the location of work (Norcross, GA; Broomfield, CO):  ").split(",")
    local_pointers["MAINT_CITY"] = location[0].strip()
    local_pointers["MAINT_STATE"] = location[1].strip()

# Primary Maintenance Date
    prim_maint_date = input("Enter the primary maintenance date in EST (YYYY-MM-DD HH:MM to YYYY-MM-DD HH:MM):  ").split(" to ")
    start_datetime = prim_maint_date[0].strip().split(" ")
    start_time = start_datetime[1].strip().split(":")

    local_pointers["MAINT_PRIM_START_DATE"] = start_datetime[0].strip()
    local_pointers["MAINT_PRIM_START_HOUR"] = start_time[0]
    local_pointers["MAINT_PRIM_START_MIN"] = start_time[1]

    end_datetime = prim_maint_date[1].strip().split(" ")
    end_time = end_datetime[1].strip().split(":")

    local_pointers["MAINT_PRIM_END_DATE"] = end_datetime[0].strip()
    local_pointers["MAINT_PRIM_END_HOUR"] = end_time[0]
    local_pointers["MAINT_PRIM_END_MIN"] = end_time[1]

# Backup Maintenance Date
    has_backup_date = input("Is there a backup night on file? If there is type 'yes':  ")
    if "yes" in has_backup_date.lower():
            back_maint_date = input("Enter the backup maintenance date in EST (YYYY/MM/DD HH:MM - YYYY/MM/DD HH:MM):  ").split("-")
            start_datetime = back_maint_date[0].strip().split(" ")
            start_time = start_datetime[1].strip().split(":")

            local_pointers["MAINT_BACK_START_DATE"] = start_datetime[0].strip()
            local_pointers["MAINT_BACK_START_HOUR"] = start_time[0]
            local_pointers["MAINT_BACK_START_MIN"] = start_time[1]

            end_datetime = back_maint_date[1].strip().split(" ")
            end_time = end_datetime[1].strip().split(":")

            local_pointers["MAINT_BACK_END_DATE"] = end_datetime[0].strip()
            local_pointers["MAINT_BACK_END_HOUR"] = end_time[0]
            local_pointers["MAINT_BACK_END_MIN"] = end_time[1]

    print("All information should have been recieved correctly. The script will run in 5 seconds. Please bring up an empty GCR shell.")
    time.sleep(6)

def _exec():
    email_tuple = extract_email_tuple()
    populate_pointers(email_tuple)