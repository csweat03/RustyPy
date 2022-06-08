import os
import json
import datetime
import time

from pynput.mouse import Controller as ControlMouse
from pynput.keyboard import Key, Controller as ControlKeybd

mouse = ControlMouse()
keybd = ControlKeybd()

parent_fldr = os.path.dirname(os.getcwd()) + "\\main\\"
config = json.loads(open(parent_fldr + "config.json", "r").read())

speed_multiplier = float(config["speed_multiplier"])

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
    'MAINT_CITY': "",
    'MAINT_STATE': "",
    'MAINT_PRIM_START_DATE': "",
    'MAINT_PRIM_END_DATE': "",
    'MAINT_PRIM_START_HOUR': "",
    'MAINT_PRIM_END_HOUR': "",
    'MAINT_PRIM_START_MIN': "",
    'MAINT_PRIM_END_MIN': "",
    # 
    'MAINT_BACK_START_DATE': "",
    'MAINT_BACK_END_DATE': "",
    'MAINT_BACK_START_HOUR': "",
    'MAINT_BACK_END_HOUR': "",
    'MAINT_BACK_START_MIN': "",
    'MAINT_BACK_END_MIN': "",
    # 'CRCT_LIST': "",
    # 'CRCT_RSLT': ""
}

categories = [
    "GCR Details",
    "Change Model",
    "Locations",
    "Windows",
    "Affected Objects",
    "Service Impacts",
    "Notifications",
    "Activity Log",
    "Documents",
    "Related Items",
    "BRUL Violations"
]

type_four = [
    "config change",
    "decommision",
    "hardware replacement",
    "hardware upgrade",
    "maintenance",
    "migration",
    "node insert",
    "software upgrade",
    "splice\\relocation",
    "troubleshoot alarm"
]

time_units = [
    "milliseconds",
    "seconds",
    "minutes",
    "hours"
]

vendors = config["vendors"]

def log(str):
    with open(log_file, 'a') as log:
        msg = "[" + datetime.datetime.now().strftime("%Y-%m-%d, %H:%M:%S") + "]: " + str + "\n"
        log.write(msg)

def press_key(key):
    try:
        keybd.press(key)
        keybd.release(key)
    except:
        log("A key has raised an exception.")
    time.sleep(0.02/(10 if speed_multiplier > 10 else 0.1 if speed_multiplier < 0.1 else speed_multiplier))

def macro_key(key, index):
    i = 0
    while i < index:
        i += 1
        press_key(key)

def macro_combo(modifier, key, index):
    with keybd.pressed(modifier):
        macro_key(key, index)

def press_tab():
    press_key(Key.tab)

def press_combo(modifier, key):
    with keybd.pressed(modifier):
        press_key(key)

def type_text(text):
    for letter in text:
        press_key(letter)

def get_keyboard_operation(array, option):
    index = array.index(option.lower())
    macro_key(Key.down, index + 1)

def wait_for_user_input(request):
    unpause = ""
    # Start Manual Process
    while (unpause.upper() != "OK"): # In future: Replace with function find call.
        unpause = input(request + ": Type OK and click on the box you were on, you have 5 seconds.")
    unpause = ""

    time.sleep(5)
    # End Manual Process

def convert_time(military):
    return int(military) // 12

def move_category(current, new):
    _from = categories.index(current)
    _to = categories.index(new)
    _delta = _to - _from
    _forwards = _delta > 0
    _magnitude = abs(_delta)
    macro_key(Key.right if _forwards else Key.left, _magnitude)


# from screeninfo import get_monitors # pip install screeninfo
#
# def find_current_monitor(x, y):
#     for display in get_monitors():
#         if (display.x > x): continue
#         if (display.x + display.width < x): continue
#         return display
#
# def build_gcr():
#     with Listener(on_click=on_click) as listener:
#         listener.join()

# def on_click(x, y, button, pressed):
#     if (pressed != True): return

#     monitor = find_current_monitor(x, y)

#     x -= monitor.x
#     y -= monitor.y

#     x = (x - monitor.x) / (monitor.width)
#     y = (y - monitor.y) / (monitor.height)

#     if (button == Button.left):
#         macro_steps.append(str(x)+","+str(y))
#     if (button == Button.right):
#         print(monitor)
#         macro_steps.append("PAUSE")
#     if (button == Button.middle):
#         print(macro_steps)
#         return False