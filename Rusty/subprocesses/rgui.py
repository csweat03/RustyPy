from ctypes.wintypes import RGB
import json
from time import sleep
from turtle import screensize
import pyautogui
import PySimpleGUI as gui

data = json.load(open('metadata.json'))

title = data["name"] + " " + data["version"] + "." + data["release"] + " - " + data["author"]

gui.theme('Reddit')

left = [
    [gui.Button("Start", tooltip="Start running the automation.", size=(10, 2))],
    [gui.Button("Stop", tooltip="Stop running the automation.", size=(10, 2))],
    [gui.Button("Pause", tooltip="Pause the app and resume where left.", size=(10, 2))],
    [gui.Button("Exit", tooltip="Close the app and abandon everything.", size=(10, 2))],
    [gui.Text("Bot Speed:", pad=(0,5))],
    [gui.Slider(range=(1,10), default_value=5, orientation="h", size=(8, 10), trough_color="#4c80d4", pad=(0,5))],
    [gui.HSeparator(pad=(0,5))],
    [gui.Button("...", tooltip="[Coming Soon] Try to pre-populate fields with AI generated input.", size=(10, 2), button_color="#56637a")]
]

right = [
    [gui.Text("Name:"), gui.Combo(['Spectrum Enterprise'], key="name", size=(20, 0)), gui.VSeperator(), gui.Text("ID:"), gui.Input(size=(20, 0))],
    [gui.Text("Description: "), gui.Input(size=(41, 0), pad=(3,0))],
    [gui.Text("Type:"), gui.Combo(['Config Change', 'Decommision', 'Hardware Replacement', 'Hardware Upgrade', 'Maintenance', 'Migration', 'Node Insert', 'Software Upgrade', 'Splice/Relocation', 'Troubleshoot Alarm'],default_value='Maintenance', pad=(2,0), key='type4'), gui.VSeperator(), gui.Text("Outage (mins):", pad=(0,0)), gui.Input(size=(10, 0), pad=(0,0))],
    [gui.Text("LOC:"), gui.Input(size=(15, 0), tooltip="City", pad=(0,0)), gui.Text(",", pad=(0,0)), gui.Input(size=(5, 0), tooltip="State Abbreviation", pad=(0,0)), gui.Text("", pad=(1,0)), gui.VSeperator(), gui.Text("Country:"), gui.Input(size=(15, 0), pad=(0,0))],
    [gui.HSeparator()],
    [gui.Text("Maint. Windows - Formats: Date = YYYY-MM-DD, Time = HH:MM", pad=(0,0))],
    [gui.Text("#1 Date:", pad=(0,2)), gui.Input(size=(10, 0), pad=(0,2)), gui.Text("Start Time:", pad=(0,2)), gui.Input(size=(5, 0), pad=(0,2)), gui.Text("End Time:", pad=(0,2)), gui.Input(size=(5, 0), pad=(0,2)), gui.Text("Primary")],
    [gui.Text("#2 Date:", pad=(0,2)), gui.Input(size=(10, 0), pad=(0,2)), gui.Text("Start Time:", pad=(0,2)), gui.Input(size=(5, 0), pad=(0,2)), gui.Text("End Time:", pad=(0,2)), gui.Input(size=(5, 0), pad=(0,2)), gui.Checkbox("Alt?")],
    [gui.Text("#3 Date:", pad=(0,2)), gui.Input(size=(10, 0), pad=(0,2)), gui.Text("Start Time:", pad=(0,2)), gui.Input(size=(5, 0), pad=(0,2)), gui.Text("End Time:", pad=(0,2)), gui.Input(size=(5, 0), pad=(0,2)), gui.Checkbox("Alt?")],
]

layout = [
    [
        gui.Column(left, element_justification='c'),
        gui.VSeperator(color="#4c80d4"),
        gui.Column(right, vertical_alignment="top"),
    ]
]

def _exec():
    width = 635
    window = gui.Window(title, layout, keep_on_top=True, element_justification='right', size=(width, 375), location=(pyautogui.size()[0] - width - 10, 45), margins=(3,3))

    window.finalize()

    while True:
        event, values = window.read()

        # window.move(window.get_screen_size()[0] - window.size[0] - 15, 45)

        # Terminate window on OK or if window closed.
        if event == "OK" or event == gui.WIN_CLOSED:
            break
