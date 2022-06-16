import time
from .utils.wrapper import *

POINTERS = internal_pointers
DETAILS = POINTERS["MAINTENANCE_DETAILS"]

def submit_date_interface(date, hour, min):
    press_combo(Key.ctrl_l, "a")
    type_text(date)
    macro_key(Key.tab, 2)
    type_text(hour) # Recieved Date
    press_key(Key.right)
    type_text(min)
    press_key(Key.right)
    press_key(Key.right)
    press_key(convert_time(hour))


def create_shell():
    press_key("c") # CMT
    time.sleep(5) # In future: Replace with function wait call, scan for status screen to get a more accurate wait time.
    press_tab()
    macro_key('s', 8) # In future: Replace with function find call, to locate the users name. This finds "Sweat, Christian" in the system.
    press_tab()
    press_key('v') # Vendor Maintenance
    press_tab()
    press_key('n') # Network
    press_tab()
    press_key('v') # Vendor
    press_tab()
    press_key('t') # Transport
    press_tab()
    get_keyboard_operation(type_four, DETAILS["category"]) # Type 4
    press_tab()
    press_key('o') # Outage
    press_key(Key.enter)
    macro_key(Key.tab, 3)
    press_key(Key.up)
    press_key(Key.enter) # Selects Change Model
    time.sleep(1)
    log("Created GCR Shell")

def submit_details():
    macro_combo(Key.shift, Key.tab, 7)
    press_key("n") # Assuming North America; Dangerous
    macro_combo(Key.shift, Key.tab, 4)
    press_key("n") # No BPMS
    macro_key(Key.tab, 18)
    press_key(Key.down) # CHNGCOORD
    press_tab()

    wait_for_user_input("Select the vendor from the Vendor Name field. ") # Vendor Name

    press_tab()
    type_text(to_string(POINTERS["IDENTIFIERS"])) # Maintenance ID
    press_tab()

    submit_date_interface(
        DETAILS["received"][0:2],
        DETAILS["received"][3],
        DETAILS["received"][4]
    )

    press_tab()
    type_text("Vendor - " + POINTERS["VENDOR_NAME"] + " - " + DETAILS["description"]) # Description
    macro_key(Key.enter, 2)
    type_text("Outage: " + DETAILS["outage"][0] + " " + DETAILS["outage"][1])
    macro_combo(Key.shift, Key.tab, 14)

def submit_location(location):
    macro_key(Key.tab, 5)
    press_key(Key.enter)
    macro_key(Key.tab, 2)
    type_text(location[0])
    press_tab()
    type_text(location[1])
    macro_key(Key.tab, 2)

    wait_for_user_input("Select the country from the Country field. ") # Country

    press_key(Key.enter)
    macro_combo(Key.shift, Key.tab, 5)

def submit_window(window):
    macro_key(Key.tab, 2)
    press_key(Key.enter)

    time.sleep(5) # In future: Replace with function wait call, scan for status screen to get a more accurate wait time.

    macro_key(Key.tab, 2)

    submit_date_interface(to_string(window[0][0][0:2]), window[0][1][0], window[0][1][1]) # Start Date

    press_tab()

    submit_date_interface(to_string(window[1][0][0:2]), window[1][1][0], window[1][1][1]) # End Date

    macro_key(Key.tab, 5) # Selecting alternate night, if neccessary.
    if window[2]:
        press_key(Key.up)
    macro_key(Key.tab, 2)

    press_key(Key.space) # Assuming only one location, and it is always active. Dangerous

    press_key(Key.enter)
    macro_combo(Key.shift, Key.tab, 2)

def submit_affected_objects(window):
    # macro_key(Key.tab, 10)
    # press_key(Key.enter)
    # press_key(Key.up)
    # time.sleep(1)
    # type_text("CMT Search - Transport and IP")
    # macro_key(Key.enter, 2)
    wait_for_user_input("Generate the Circuit IDs for Maintenace Window #" + (DETAILS.index(window) + 1)) # Country

    macro_key(Key.tab, 2)
    press_key(Key.enter)
    time.sleep(1)
    press_key("v") # Vendor
    press_tab()
    press_key("c") # CenturyLink
    macro_key(Key.tab, (DETAILS.index(window) + 1))
    press_key(Key.space) # Assuming only one maint. window, and it is always active. Dangerous
    macro_key(Key.tab, 2) # User will now be prompted to copy the circuits for night 1, 2 and 3 seperately ^
    type_text(DETAILS["outage"][0])
    press_tab()
    get_keyboard_operation(time_units, DETAILS["outage"][1])
    press_tab()
    type_text(DETAILS["outage"][0])
    press_tab()
    get_keyboard_operation(time_units, DETAILS["outage"][1])
    press_tab()
    press_key(Key.enter)
    time.sleep(1)
    press_combo(Key.shift, Key.tab)
    press_combo(Key.ctrl, "v")
    macro_combo(Key.ctrl, Key.tab, 2)
    press_key(Key.enter)
    time.sleep(1)
    press_key(Key.enter)
    time.sleep(3)


def _exec():
    create_shell()
    submit_details()
    move_category("GCR Details", "Locations")
    submit_location(DETAILS["location"])
    move_category("Locations", "Windows")
    for window in DETAILS["windows"]:
        submit_window(window)
    move_category("Locations", "Windows")
    for window in DETAILS["windows"]:
        submit_affected_objects(window)

    # macro_key(Key.tab, 8)
    # press_key(Key.enter)
    # time.sleep(5)
