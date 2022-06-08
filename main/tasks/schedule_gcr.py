import time
from .utils.wrapper import *

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
    get_keyboard_operation(type_four, local_pointers["MAINT_TYPE"]) # Type 4
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
    type_text(local_pointers["MAINT_ID"]) # Maintenance ID
    press_tab()

    submit_date_interface(
        local_pointers["MAINT_DATE_REC"],
        local_pointers["MAINT_HOUR_REC"],
        local_pointers["MAINT_MIN_REC"]
    )

    press_tab()
    type_text("Vendor - " + local_pointers["VNDR_NAME"] + " - " + local_pointers["MAINT_DESC"]) # Description
    macro_key(Key.enter, 2)
    type_text("Outage: " + local_pointers["MAINT_OUTAGE_NUM"] + " " + local_pointers["MAINT_OUTAGE_UNT"])
    macro_combo(Key.shift, Key.tab, 14)

def submit_location(city, state):
    macro_key(Key.tab, 5)
    press_key(Key.enter)
    macro_key(Key.tab, 2)
    type_text(city)
    press_tab()
    type_text(state)
    macro_key(Key.tab, 2)

    wait_for_user_input("Select the country from the Country field. ") # Country

    press_key(Key.enter)
    macro_combo(Key.shift, Key.tab, 5)

def submit_window(start_date, start_hour, start_min, end_date, end_hour, end_min, alternate_night):
    macro_key(Key.tab, 2)
    press_key(Key.enter)

    time.sleep(5) # In future: Replace with function wait call, scan for status screen to get a more accurate wait time.

    macro_key(Key.tab, 2)

    submit_date_interface(start_date, start_hour, start_min) # Start Date

    press_tab()

    submit_date_interface(end_date, end_hour, end_min) # End Date

    macro_key(Key.tab, 5) # Selecting alternate night, if neccessary.
    if alternate_night:
        press_key(Key.up)
    macro_key(Key.tab, 2)

    press_key(Key.space) # Assuming only one location, and it is always active. Dangerous

    press_key(Key.enter)
    macro_combo(Key.shift, Key.tab, 2)

def submit_affected_objects():
    # macro_key(Key.tab, 10)
    # press_key(Key.enter)
    # press_key(Key.up)
    # time.sleep(1)
    # type_text("CMT Search - Transport and IP")
    # macro_key(Key.enter, 2)
    wait_for_user_input("Please generate the GCR Impact, Copy the list to your clipboard, and click the Affected Object category.") # Country

    macro_key(Key.tab, 2)
    press_key(Key.enter)
    time.sleep(1)
    press_key("v") # Vendor
    press_tab()
    press_key("c") # CenturyLink
    press_tab()
    press_key(Key.space) # Assuming only one maint. window, and it is always active. Dangerous
    macro_key(Key.tab, 2)
    type_text(local_pointers["MAINT_OUTAGE_NUM"])
    press_tab()
    get_keyboard_operation(time_units, local_pointers["MAINT_OUTAGE_UNT"])
    press_tab()
    type_text(local_pointers["MAINT_OUTAGE_NUM"])
    press_tab()
    get_keyboard_operation(time_units, local_pointers["MAINT_OUTAGE_UNT"])
    press_tab()
    press_key(Key.enter)
    time.sleep(1)
    press_combo(Key.shift, Key.tab)
    press_combo(Key.ctrl, "v")
    macro_combo(Key.ctrl, Key.tab, 2)
    press_tab()
    press_key(Key.enter)
    macro_combo(Key.shift, Key.tab, 2)
    time.sleep(5)


def _exec():
    create_shell()
    submit_details()
    move_category("GCR Details", "Locations")
    submit_location(local_pointers["MAINT_CITY"], local_pointers["MAINT_STATE"])
    move_category("Locations", "Windows")
    submit_window(
        local_pointers["MAINT_PRIM_START_DATE"],
        local_pointers["MAINT_PRIM_START_HOUR"],
        local_pointers["MAINT_PRIM_START_MIN"],
        local_pointers["MAINT_PRIM_END_DATE"],
        local_pointers["MAINT_PRIM_END_HOUR"],
        local_pointers["MAINT_PRIM_END_MIN"],
        False
    )
    if (len(local_pointers["MAINT_BACK_START_DATE"]) > 5):
        submit_window(
        local_pointers["MAINT_BACK_START_DATE"],
        local_pointers["MAINT_BACK_START_HOUR"],
        local_pointers["MAINT_BACK_START_MIN"],
        local_pointers["MAINT_BACK_END_DATE"],
        local_pointers["MAINT_BACK_END_HOUR"],
        local_pointers["MAINT_BACK_END_MIN"],
        True
    )
    move_category("Locations", "Windows")
    submit_affected_objects()

    macro_key(Key.tab, 8)
    press_key(Key.enter)
    time.sleep(5)
