import os
import json
import datetime
import time
import win32com.client # pip install -U pypiwin32

from screeninfo import get_monitors # pip install screeninfo
from pynput.mouse import Button, Controller as ControlMouse, Listener
from pynput.keyboard import Key, Controller as ControlKeybd

parent_fldr = os.path.dirname(os.getcwd()) + "\\"
config = json.loads(open(parent_fldr + "config.json", "r").read())

input_fldr = parent_fldr + config["input_fldr"]
log_file = parent_fldr + config["log"]
pointers = parent_fldr + config["pointers"]

email_tuple = ()
local_pointers = {
    'VNDR_NAME': "",
    'MAINT_ID': "",
    'MAINT_TYPE': "",
    'MAINT_DESC': "",
    'MAINT_DATE_REC': "",
    'MAINT_HOUR_REC': "",
    'MAINT_MIN_REC': "",
    'MAINT_OUTAGE_NUM': "",
    'MAINT_OUTAGE_UNT': "",
    # 'MAINT_CITY': "",
    # 'MAINT_STATE': "",
    # 'MAINT_PRIM_START_DATE': "",
    # 'MAINT_PRIM_END_DATE': "",
    # 'MAINT_PRIM_START_HOUR': "",
    # 'MAINT_PRIM_END_HOUR': "",
    # 'MAINT_PRIM_START_MIN': "",
    # 'MAINT_PRIM_END_MIN': "",
    # 
    # 'MAINT_BACK_START_DATE': "",
    # 'MAINT_BACK_END_DATE': "",
    # 'MAINT_BACK_START_HOUR': "",
    # 'MAINT_BACK_END_HOUR': "",
    # 'MAINT_BACK_START_MIN': "",
    # 'MAINT_BACK_END_MIN': "",
    # 'CRCT_LIST': "",
    # 'CRCT_RSLT': ""
}

type_four = {
    "config change": 1,
    "decommision": 2,
    "hardware replacement": 3,
    "hardware upgrade": 4,
    "maintenance": 5,
    "migration": 6,
    "node insert": 7,
    "software upgrade": 8,
    "splice\\relocation": 9,
    "troubleshoot alarm": 10
}

display_list = []
macro_steps = []

mouse = ControlMouse()
keybd = ControlKeybd()

def log(str):
    with open(log_file, 'a') as log:
        msg = "[" + datetime.datetime.now().strftime("%m/%d/%Y, %H:%M:%S") + "]: " + str + "\n"
        log.write(msg)
        
def read_message(absolute_file):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    message = outlook.OpenSharedItem(absolute_file)
    return message

def extract_email_tuple():
    for file in os.scandir(input_fldr):
        f = file.name

        log("Found file: " + f)
        msg = read_message(input_fldr + "\\" + file.name)

        mail_sent_on = msg.SentOn.strftime("%m/%d/%Y, %H:%M:%S")
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
    #temporary manual entry, would like to replace with AI/ML in the future
    for pointer in local_pointers:
        local_pointers[pointer] = input("Enter this pointer's value " + pointer + ":  ")

    print("All information should have been recieved correctly. The script will run in 5 seconds. Please bring up an empty GCR shell.")
    time.sleep(6)

def press_key(key):
    keybd.press(key)
    keybd.release(key)

def macro_key(key, index):
    i = 0
    while i < index:
        i += 1
        press_key(key)

def press_tab():
    press_key(Key.tab)

def type_text(text):
    keybd.type(text)

def get_keyboard_operation(input, option_list):
    index = option_list[str(input).lower()]
    macro_key(Key.down, index)

def build_gcr():
    with Listener(on_click=on_click) as listener:
        listener.join()

def find_monitor(x, y):
    for display in display_list:
        if (display.x > x): continue
        if (display.x + display.width < x): continue
        return display

def on_click(x, y, button, pressed):
    if (pressed != True): return

    monitor = find_monitor(x, y)

    x -= monitor.x
    y -= monitor.y

    x = (x - monitor.x) / (monitor.width)
    y = (y - monitor.y) / (monitor.height)

    if (button == Button.left):
        macro_steps.append(str(x)+","+str(y))
    if (button == Button.right):
        print(monitor)
        macro_steps.append("PAUSE")
    if (button == Button.middle):
        print(macro_steps)
        return False


running = "no"
unpause = ""

while running != "run":
    running = input("Put all email files inside the folder labeled \" \\input\\ \", type RUN and press ENTER when done:  ").lower()

for m in get_monitors():
        display_list.append(m)
    
# email_tuple = extract_email_tuple()
populate_pointers(None)
press_key("c")
time.sleep(3)
press_tab()
macro_key('s', 8)
press_tab()
press_key('v')
press_tab()
press_key('n')
press_tab()
press_key('v')
press_tab()
press_key('t')
press_tab()
get_keyboard_operation(local_pointers["MAINT_TYPE"], type_four)
press_tab()
press_key('o')
press_key(Key.enter)
macro_key(Key.tab, 3)
press_key(Key.up)
press_key(Key.enter)
time.sleep(1)
macro_key(Key.tab, 8)
press_key(Key.down)
press_tab()

while (unpause.upper() != "OK"):
    unpause = input("Select the vendor from the Vendor Name field and type OK when done: ")
unpause = ""

time.sleep(3)

press_tab()
type_text(local_pointers["MAINT_ID"])
press_tab()
type_text(local_pointers["MAINT_DATE_REC"])
macro_key(Key.tab, 2)
type_text(local_pointers["MAINT_HOUR_REC"])
ampm = "P" if int(local_pointers["MAINT_HOUR_REC"]) >= 12 else "A"
press_key(Key.right)
type_text(local_pointers["MAINT_MIN_REC"])
press_key(Key.right)
press_key(Key.right)
press_key(ampm)
press_tab()
type_text("Vendor - " + local_pointers["VNDR_NAME"] + " - " + local_pointers["MAINT_DESC"])
macro_key(Key.enter, 2)
type_text("Outage: " + local_pointers["MAINT_OUTAGE_NUM"] + " " + local_pointers["MAINT_OUTAGE_UNT"])
with keybd.pressed(Key.shift):
    macro_key(Key.tab, 14)
macro_key(Key.right, 2)
# macro_key(Key.tab, 5)
# press_key(Key.enter)
# macro_key(Key.tab, 2)
# type_text(local_pointers["MAINT_CITY"])
# press_tab()
# type_text(local_pointers["MAINT_STATE"])
# macro_key(Key.tab, 2)

# while (unpause.upper() != "OK"):
#     unpause = input("Select the country from the Country field and type OK when done: ")
# unpause = ""

# time.sleep(3)

# press_key(Key.enter)
# with keybd.pressed(Key.shift):
#     macro_key(Key.tab, 5)
# macro_key(Key.right, 1)
# macro_key(Key.tab, 2)
# press_key(Key.enter)

# time.sleep(3)

# macro_key(Key.tab, 2)
# with keybd.pressed(Key.ctrl_l):
#     press_key("a")
# type_text(local_pointers["MAINT_PRIM_START_DATE"])
# macro_key(Key.tab, 2)
# type_text(local_pointers["MAINT_PRIM_START_HOUR"])
# ampm = "P" if int(local_pointers["MAINT_PRIM_START_HOUR"]) >= 12 else "A"
# press_key(Key.right)
# type_text(local_pointers["MAINT_PRIM_START_MIN"])
# press_key(Key.right)
# press_key(Key.right)
# press_key(ampm)
# press_tab()
# # 
# with keybd.pressed(Key.ctrl_l):
#     press_key("a")
# type_text(local_pointers["MAINT_PRIM_END_DATE"])
# macro_key(Key.tab, 2)
# type_text(local_pointers["MAINT_PRIM_END_HOUR"])
# ampm = "P" if int(local_pointers["MAINT_PRIM_END_HOUR"]) >= 12 else "A"
# press_key(Key.right)
# type_text(local_pointers["MAINT_PRIM_END_MIN"])
# press_key(Key.right)
# press_key(Key.right)
# press_key(ampm)
# macro_key(Key.tab, 7)
# press_key(Key.space)
# press_key(Key.enter)

# if (len(local_pointers["MAINT_BACK_START_DATE"]) > 5):
#     press_key(Key.enter)
#     time.sleep(3)

#     macro_key(Key.tab, 2)
#     with keybd.pressed(Key.ctrl_l):
#         press_key("a")
#     type_text(local_pointers["MAINT_BACK_START_DATE"])
#     macro_key(Key.tab, 2)
#     type_text(local_pointers["MAINT_BACK_START_HOUR"])
#     ampm = "P" if int(local_pointers["MAINT_BACK_START_HOUR"]) >= 12 else "A"
#     press_key(Key.right)
#     type_text(local_pointers["MAINT_BACK_START_MIN"])
#     press_key(Key.right)
#     press_key(Key.right)
#     press_key(ampm)
#     press_tab()
#     # 
#     with keybd.pressed(Key.ctrl_l):
#         press_key("a")
#     type_text(local_pointers["MAINT_BACK_END_DATE"])
#     macro_key(Key.tab, 2)
#     type_text(local_pointers["MAINT_BACK_END_HOUR"])
#     ampm = "P" if int(local_pointers["MAINT_BACK_END_HOUR"]) >= 12 else "A"
#     press_key(Key.right)
#     type_text(local_pointers["MAINT_BACK_END_MIN"])
#     press_key(Key.right)
#     press_key(Key.right)
#     press_key(ampm)
#     macro_key(Key.tab, 5)
#     press_key(Key.up)
#     macro_key(Key.tab, 2)
#     press_key(Key.space)
#     press_key(Key.enter)




# build_gcr()