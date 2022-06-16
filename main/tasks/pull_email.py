import math
from operator import contains
import os
import datetime
import time
import re
from .utils.wrapper import *
import win32com.client # pip install -U pypiwin32
        
def read_message(absolute_file):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    message = outlook.OpenSharedItem(absolute_file)
    return message

def extract_window(window_index, alternate_night, seperator_list):
    maint_part = None

    while True:
        prim_maint_date = input(("(ALT) " if alternate_night else "") + "Enter Maint. Date #" + str(window_index) + " in EST (Format: YYYY-MM-DD HH:MM to YYYY-MM-DD HH:MM):  ")
        try:
            maint_part = prim_maint_date.split("to")
        except:
            maint_part = prim_maint_date.split(" - ")
        finally:
            if len(maint_part) >= 2:
                break
        print("ERROR: Could not decipher the date provided.. Try this format exactly: YYYY-MM-DD HH:MM to YYYY-MM-DD HH:MM")
        
    start_datetime = maint_part[0].strip().split()
    start_date = re.split(seperator_list, start_datetime[0].strip())
    start_time = re.split(seperator_list, start_datetime[1].strip())

    end_datetime = maint_part[1].strip().split()
    end_date = re.split(seperator_list, end_datetime[0].strip())
    end_time = re.split(seperator_list, end_datetime[1].strip())
    # ((("", "", ""), ("", "")), (("", "", ""), ("", "")), False)

    internal_pointers["MAINTENANCE_DETAILS"]["windows"].append((
        ((start_date[0], start_date[1], start_date[2]), (start_time[0], start_time[1])),
        ((end_date[0], end_date[1], end_date[2]), (end_time[0], end_time[1])), 
        alternate_night))


def extract_email_tuple():
    for file in os.scandir(input_fldr):
        f = file.name

        log("Found file: " + f)
        msg = read_message(input_fldr + "\\" + file.name)

        mail_sent_on = msg.SentOn.strftime("%Y/%m/%d/%H/%M/%S")
        mail_content = msg.Subject.strip()

        for line in msg.Body.strip().split("\n"):
            line = line.strip()
            if (line):
                mail_content += " " + line

        log("!----- Recieved: \"" + mail_sent_on + "\"")
        log("!----- Contents: \"" + mail_content + "\"")

        # for word in mail_content:
        #     if find_dates(word) != None:
        #         print(str(valid_dates))

        # print("\n"*5)

        return (mail_sent_on, mail_content)

def grab_vendor_name(email_tuple):
    vendor_list = []
    # Vendor Name
    for vendor in vendors.split("|"):
        if email_tuple[1].lower().__contains__(vendor.lower()) and not vendor in vendor_list:
            vendor_list.append(vendor)

    if len(vendor_list) <= 1:
        internal_pointers["VENDOR_NAME"] = input('Enter the vendor name:  ')
        write("vendor_list", internal_pointers["VENDOR_NAME"])
    else:
        log("Found Vendor Names in Mail Subject: " + to_string(vendor_list))
        print("Vendor Name: " + to_string(vendor_list))
        internal_pointers["VENDOR_NAME"] = vendor_list[0]

def grab_identifiers(email_tuple):
    ID = internal_pointers["IDENTIFIERS"]

    for id in identifiers.split("|"):
        reg_id = '('
        for char in id:
            if char == '?':
                reg_id += '[0-9]'
            else:
                reg_id += '[' + char + ']'
        reg_id += ')+'

        result = re.findall(reg_id, email_tuple[1])

        for ret in result:
            if ret not in ID:
                ID.append(ret)

    if len(ID) == 0:
        ID.append(input("Enter the maintenance ID:  "))
        clone = []
        for id in ID:
            clone.append(re.sub('[^a-zA-Z]', '?', id))
        write("identifier_list", clone[0])

    if len(ID) > 0:
        log("Found the following Maintenance IDs in E-mail:\n\t" + to_string(ID))
        print("Maintenance IDs: " + to_string(ID))

def grab_received_date(email_tuple, seperator_list, details):
    details["received"] = re.split(seperator_list, email_tuple[0])
    log("Found the Date Recieved in E-mail:\n\t" + to_string(details["received"]))
    print("Date Recieved has been found.")

def grab_description_with_type(email_tuple, details):
    listening = 0
    triggers = ["description", "comments", "reason"]
    description = ""

    threshold = 0.9
    passes = 3

    for word in email_tuple[1].split(" "):
        if "http" in word:
            continue
        for trigger in triggers:
            if calculate_string_similarity(trigger.lower(), word.lower()) >= threshold:
                listening = passes
                continue

        if listening > 0:
            description += word + " "
            if word == "\n" or word == "\t" or word == "  ":
                listening = listening - 1

    raw = description.replace("\t", " ").replace("\n", " ").strip().replace("  ", " ")

    vals = [
        0.0,
        0.0,
        0.0,
        0.0,
        0.0,
        0.0,
        0.0,
        0.0,
        0.0,
        0.0
    ]

    for cat in type_four:
        for word in cat.split(" "):
            for syn in get_synonym_list(word):
                if syn in raw:
                    pres = (calculate_string_similarity(word, syn))
                    past = vals[type_four.index(cat)]
                    vals[type_four.index(cat)] = min((past + pres) / 2, 1)
    
    probable = type_four[vals.index(max(vals))] if max(vals) > 0.99 else type_four[4]

    details["category"] = probable
    print("Best Suited Category: " + details["category"])
    
    details["description"] = summarize_text(raw, 1.25)
    print("Partial Description: " + details["description"])

def populate_pointers(email_tuple):
    seperator_list = '/-_:;|\\<>,.[ ]'
    
    grab_vendor_name(email_tuple)
    grab_identifiers(email_tuple)

    DETAILS = internal_pointers["MAINTENANCE_DETAILS"]

    grab_received_date(email_tuple, seperator_list, DETAILS)

    grab_description_with_type(email_tuple, DETAILS)

# Outage Duration
    outage = input("Enter the outage duration (6 hours, 30 minutes, etc):  ").split(" ")
    DETAILS["outage"].append(outage[0])
    DETAILS["outage"].append(outage[1])

# Work Location
    location = input("Enter the location of work (Norcross, GA; Broomfield, CO):  ").split(",")
    DETAILS["location"].append(location[0].strip())
    DETAILS["location"].append(location[1].strip())



# Primary and Backup Maintenance Dates
    num_windows = int(input("How many primary windows are there? (1, 2, 3):  "))
    i = 1
    while True:
        extract_window(i, False, seperator_list)
        i += 1
        if i > num_windows:
            break

    has_alt_night = input("Is there an alternate night?  ")

    if has_alt_night: # Might have to set to i+1, we'll see :)
        extract_window(i, True, seperator_list)

    print("All information should have been recieved correctly. The script will run in 5 seconds. Please bring up an empty GCR shell.")
    time.sleep(6)

def _exec():
    email_tuple = extract_email_tuple()
    populate_pointers(email_tuple)