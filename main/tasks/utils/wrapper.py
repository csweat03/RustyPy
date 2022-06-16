import os
import json
import datetime
import time

import dateparser

import nltk
from nltk.corpus import stopwords, wordnet
from nltk.tokenize import word_tokenize, sent_tokenize
nltk.download('stopwords')
nltk.download('omw-1.4')
nltk.download('punkt')
print("\n")

from pynput.mouse import Controller as ControlMouse
from pynput.keyboard import Key, Controller as ControlKeybd
from difflib import SequenceMatcher

mouse = ControlMouse()
keybd = ControlKeybd()

parent_fldr = os.path.dirname(os.getcwd()) + "\\main\\"
config = json.loads(open(parent_fldr + "config.json", "r").read())

speed_multiplier = float(config["speed_multiplier"])

input_fldr = parent_fldr + config["input_fldr"]
storage_fldr = parent_fldr + config["storage_fldr"]

pointers = parent_fldr + config["pointers"]

email_tuple = ()

internal_pointers = {
    "VENDOR_NAME": "",
    "IDENTIFIERS": [],
    "MAINTENANCE_DETAILS": {
        # Auto-populated from mail
        "received": [],

        "description": "",
        "category": "",
        # duration, units
        "outage": [],

        "location": [],
        # window dates, window times, alternate night
        "windows": [],
    }
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

def get_synonym_list(base):
    syn = []
    for pot in wordnet.synsets(base):
        for i in pot.lemmas():
            syn.append(i.name())

    return syn

# Completely ripped from GeeksForGeeks
def summarize_text(input, magnitude):
    stopWords = set(stopwords.words("english"))
    words = word_tokenize(input)

    freqTable = dict()
    for word in words:
        word = word.lower()
        if word in stopWords:
            continue
        if word in freqTable:
            freqTable[word] += 1
        else:
            freqTable[word] = 1

    sentences = sent_tokenize(input)
    sentenceValue = dict()
   
    for sentence in sentences:
        for word, freq in freqTable.items():
            if word in sentence.lower():
                if sentence in sentenceValue:
                    sentenceValue[sentence] += freq
                else:
                    sentenceValue[sentence] = freq

    sumValues = 0
    for sentence in sentenceValue:
        sumValues += sentenceValue[sentence]

    average = int(sumValues / len(sentenceValue))

    summary = ''
    for sentence in sentences:
        if (sentence in sentenceValue) and (sentenceValue[sentence] > (magnitude * average)):
            summary += " " + sentence

    return summary


def calculate_string_similarity(a, b):
    return SequenceMatcher(None, a, b).ratio()


def read(pointer_name):
    content = ""
    point = storage_fldr + "\\" + config[pointer_name]
    with open(point, 'r') as file:
        for line in file:
            content += line.replace("\n", "|")
    return content

def write(pointer_name, content):
    content = content.strip()
    point = storage_fldr + "\\" + config[pointer_name]
    with open(point, 'a') as file:
        log("Added " + content + " to database.")
        file.write(content + "\n")


vendors = read("vendor_list")
identifiers = read("identifier_list")

valid_dates = []

def find_dates(str):
    raw = dateparser.parse(str, date_formats=['%y-%m-%d %H:%M:%S %Z'])
    if raw == None: return None
    raw_tt = raw.timetuple()
    now_tt = datetime.date.today().timetuple()
    if raw_tt.tm_yday <= now_tt.tm_yday or ((raw_tt.tm_yday - now_tt.tm_yday) + 365) % 365 > 21 or valid_dates.__contains__(raw):
        return None

    valid_dates.append(raw)
    

def to_string(l):
    return ('%-2s ' * len(l))[:-1] % tuple(l)

def log(str):
    with open(parent_fldr + config["log"], 'a') as log:
        msg = "[" + datetime.datetime.now().strftime("%Y-%m-%d, %H:%M:%S") + "]: " + str + "\n"
        try:
            log.write(msg)
        except:
            log.write("?")

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