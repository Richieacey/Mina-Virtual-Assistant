import urllib.error
import urllib.request
import pyttsx3 as py
import webbrowser
import time
import PySimpleGUI as sG
import os
import datetime
import win32api
import random
import pyaudio
import pyautogui as pag
from win32com.client import Dispatch
import wolframalpha
import winshell
import pyjokes
import subprocess
import requests
from bs4 import BeautifulSoup
import json
import keyboard as kb
from threading import Timer
import psutil
import bluetooth
import pickle
import openai
import numpy as np
from textwrap3 import wrap

# import pyi_splash
# os.startfile(r'.\SplashScreen\SplashScreen.exe')
# Update the text on the splash screen

OPENAI_API_KEY = "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"
openai.api_key = OPENAI_API_KEY


def generate_response(prompt):
    completions = openai.Completion.create(
        engine="text-davinci-002",
        prompt=prompt,
        max_tokens=1024,
        n=1,
        stop=None,
        temperature=0.5,
    )

    message = completions.choices[0].text
    return message
engine = py.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)
new_vol = 1 #volume between 0 and 1
engine.setProperty("volume", new_vol)
Name = win32api.GetComputerName()
location = 'New York'
rate = engine.getProperty('rate')
engine.setProperty('rate', 180)


def speak(audio1):
    engine.say(audio1)
    engine.runAndWait()


if not os.path.exists('encodename.pickle'):
    speak('welcome, please register your profile')
    os.startfile(r".\Profile Registry.exe")

if os.path.exists('encodename.pickle'):
    Name = pickle.load(open('encodename.pickle', 'rb'))

if os.path.exists('encodecity.pickle'):
    location = pickle.load(open('encodecity.pickle', 'rb'))


# this window is intialized when the bluetooth scanner button is pressed
# it
def open_scanner():

    layout = [[sG.Text("Device: ", key='addr', background_color="blue")], [sG.InputText(key='my-key', enable_events=True,
                             text_color="#00baff",
                             font="Arial", size=(20, 10), tooltip="enter a command", pad=((60, 0), (5, 15))),
                sG.Button("get address", key="getadd", button_color="#9855d7",
                          auto_size_button=True,
                          border_width=0, pad=((12, 15), (0, 12))),
               sG.Button("OK", button_color="#9855d7",
                         auto_size_button=True,
                         border_width=0, pad=((12, 15), (0, 12)))]]

    window = sG.Window('Bluetooth Scan', layout, background_color="blue", icon=r".\Pictures\MINA.ico",
                       titlebar_background_color="orange",
                       titlebar_font="", titlebar_text_color="white", use_custom_titlebar=True,
                       titlebar_icon=r".\Pictures\MINA-titlebar.png",
                       finalize=True,modal=True)




    while True:
        event, values = window.read()

        if event == 'getadd':
            print("scanning for devices")
            nearby_devices = bluetooth.discover_devices(duration=4, lookup_names=True,
                                                        flush_cache=True, lookup_class=False)
            number_of_devices = len(nearby_devices)
            print(number_of_devices, "devices found")
            for addr, name in nearby_devices:
                if name == "MINA HOME":
                    print("/n")
                    print(f'Device name: {name}')
                    print(f'Device address: {addr}')
                    print('/n')
                    window['addr'].update(f'Device: {name}')
                    window['my-key'].update(value=addr)
        elif event == 'OK':
            file = open(r'./bluetoothaddress.pickle', 'wb')
            address = values['my-key']
            # dump encoding in pickle file
            pickle.dump(address, file)
            file.close()
            print('done')
        elif event in (sG.WIN_CLOSED, 'Exit'):
            break
    window.close()

# checks the time of the day and wishes user accordingly
def wishme():
    hour = int(datetime.datetime.now().hour)
    if 0 <= hour < 12:
        speak(f"Good Morning {Name}")
    elif 12 <= hour < 18:
        speak(f"Good Afternoon {Name}")
    else:
        speak(f"Good Evening {Name}")


#  time function for the chat window
def asked_time():
    now = datetime.datetime.now()
    dts = now.strftime('%I:%M')
    current_time = dts
    hour = int(datetime.datetime.now().hour)
    if 0 <= hour < 12:
        speak(f"the time is {current_time + ' AM'}")
    else:
        speak(f"the time is {current_time + ' PM'}")

def play_music():
    music_dir2 = f'C:\\Users\\{Name}\\Music\\Playlists'
    songs = os.listdir(music_dir2)
    playlist = random.choice(songs)
    my_str = f'{playlist}'
    size = len(my_str)
    final_str = my_str[:size - 3]
    speak(f'playing {final_str}')
    os.startfile(os.path.join(music_dir2, playlist))


# pyi_splash.close()

def check_internet(host="https://www.google.com/"):
    try:
        urllib.request.urlopen(host)
        return True
    except Exception:
        return False


wishme()


layout1 = [[sG.Button(key="Commands",
                      image_filename=r".\Pictures\MINA.png", image_size=(120, 120),
                      button_color="#00baff", border_width=0, tooltip="click for list of commands"),
            sG.Text("---Hey there! I'm Mina. A virtual assistant\n for windows.", key="text1",
                    background_color="#00baff")],
           [sG.Checkbox("switch to question mode", background_color="#00baff", enable_events=True,key='mode'),
            sG.Button("Bluetooth Scanner", key="scan", pad=((50, 5), (5, 40)), button_color="blue", border_width=0),
            sG.Button("Connect", key="con", pad=((10, 5), (5, 40)), button_color="blue", border_width=0)],
           [sG.Button(key='btnkey', image_filename=r".\Pictures\recog-mic2.png",
                      image_size=(18, 22),
                      border_width=0, button_color="#00baff", pad=((55, 0), (5, 15))),
            sG.InputText(key='my-key', enable_events=True,
                         text_color="#00baff",
                         font="Arial", size=(20, 10), tooltip="enter a command", pad=((60, 0), (5, 15))),
            sG.Button("OK", button_color="#9855d7",
                      auto_size_button=True,
                      border_width=0, pad=((12, 15), (0, 12)))],
           [sG.Button("EXIT", button_color="red", border_width=0, pad=((50, 20), (5, 20))),
            sG.InputText(key='my-key2', enable_events=True,
                         text_color="#00baff", size=(30, 14),
                         tooltip="search google", pad=((12, 20), (5, 20))),
            sG.Button("SEARCH", button_color="#9855d7",
                      auto_size_button=True, border_width=0, pad=((2, 20), (0, 17)))],
           [sG.Button('UNREGISTER PROFILE', button_color="#9855d7",
                      auto_size_button=True, border_width=0,key='unrgstr', pad=((2, 20), (20, 17)))]]


layout2 = [[sG.Image(source=r".\Pictures\R1.png", pad=((0, 0), (20, 0)), background_color='white', key='myprofile'),
            sG.Text("New Window", background_color='white', text_color='white', key="new1", pad=((0, 170), (0, 0)))],
           [sG.Text("New Window", background_color='white', auto_size_text=True, size=(60, 5), text_color='white', key="new2",
                    pad=((170, 0), (0, 0))),
            sG.Image(source=r".\Pictures\mina81.png", background_color='white', key='mina1')],
           [sG.Image(source=r".\Pictures\R1.png", background_color='white', key='myprofile2'),
            sG.Text("New Window", background_color='white', text_color='white', key="new3", pad=((0, 170), (0, 0)))],
           [sG.Text("New Window", background_color='white', text_color='white',size=(60, 5), key="new4",
                    pad=((170, 0), (0, 0))),
            sG.Image(source=r".\Pictures\mina81.png", background_color='white', key='mina2')],
           [sG.Image(source=r".\Pictures\R1.png", background_color='white', key='myprofile3'),
            sG.Text("New Window", background_color='white', text_color='white', key="new5", pad=((0, 170), (0, 0)))],
           [sG.Text("New Window", background_color='white', text_color='white',size=(60, 5),
                    key="new6", pad=((170, 0), (0, 0))),
            sG.Image(source=r".\Pictures\mina81.png", background_color='white', key='mina3')],
           [sG.Image(source=r".\Pictures\R1.png", background_color='white', key='myprofile4'),
            sG.Text("New Window", background_color='white', text_color='white', key="new6", pad=((0, 170), (0, 0)))],
           [sG.Text("New Window", background_color='white', text_color='white',size=(60, 5),
                    key="new7", pad=((170, 0), (0, 0))),
            sG.Image(source=r".\Pictures\mina81.png", background_color='white', key='mina4')],
           [sG.InputText(key='my-key3', enable_events=True,
                         text_color="#00baff", size=(50, 10),
                         tooltip="search google"),
            sG.Button("ENTER", button_color="#9855d7",
                      auto_size_button=True,
                      border_width=0, pad=((12, 15), (0, 12)))]]

layout = [[sG.Column(layout1, element_justification='left', background_color='#00baff'),
           sG.Column(layout2, element_justification='c', pad=((12, 0), (25, 0)), background_color="white")]]

window = sG.Window("MINA", layout, resizable=True, size=(450, 320), background_color="#00baff",
                   icon=r".\Pictures\MINA.ico", titlebar_background_color="#9855d7",
                   titlebar_font="", titlebar_text_color="white", use_custom_titlebar=True,
                   titlebar_icon=r".\Pictures\MINA-titlebar.png",
                   finalize=True, modal=False)


window.bind('<Configure>', "Configure")

# myprofile1 = window['new1'].DisplayText
# mina1 = window['new2'].DisplayText
# myprofile2 = window['new3'].DisplayText
# mina2 = window['new4'].DisplayText
# myprofile3 = window['new5'].DisplayText
# mina3 = window['new6'].DisplayText

now = datetime.datetime.now()
dts = now.strftime('%I:%M')
current_time = dts


# resets the chat window
def clearchat():
    window['myprofile'].update(source=r".\Pictures\R1.png")
    window['new1'].update(value='New Window', text_color='white')
    window['myprofile2'].update(source=r".\Pictures\R1.png")
    window['new2'].update(value='New Window', text_color='white')
    window['myprofile3'].update(source=r".\Pictures\R1.png")
    window['new3'].update(value='New Window', text_color='white')
    window['new4'].update(value='New Window', text_color='white')
    window['mina1'].update(source=r".\Pictures\mina81.png")
    window['mina2'].update(source=r".\Pictures\mina81.png")
    window['mina3'].update(source=r".\Pictures\mina81.png")
    window['new5'].update(value='New Window', text_color='white')
    window['new6'].update(value='New Window', text_color='white')


valid_convo_starters = ['joke', "what's up", "how are you", "maker", "creator", "music", 'clear',
                        'fine', 'girlfriend', 'bored', 'ok', 'yes', 'alright',
                        'no', 'hello mina', 'hi mina', 'hey mina',
                        'weather', 'mina', '1', '2',
                        '3', '4', '5', '6',
                        '7', '8', '9', '0', '*', '+', '/', '-',
                        'lights on', 'lights out',
                        'desk lamp on', 'desk lamp off', 'evaluate',
                        'two minutes', 'awesome', 'great']


lightstate = ['off']
def chatbot():
    if "what's up" in values['my-key3'].lower() and "what's up" in valid_convo_starters:

        def answer():
            stmsgs = ['as of now i\'m only helping you with whatever you need me for',
                      'I am fine!, lets talk about you, are you fine?', 'oh, nothing',
                      'i am okay ! How are you?']
            ans_q = random.choice(stmsgs)
            if "what's up" in window['new1'].DisplayText.lower() and window['new2'].DisplayText == 'New Window':
                window['mina1'].update(source=r".\Pictures\mina80.png")
                window['new2'].update(value=ans_q, text_color='#9855d7')
            elif "what's up" in window['new3'].DisplayText.lower() and window['new4'].DisplayText == 'New Window':
                window['mina2'].update(source=r".\Pictures\mina80.png")
                window['new4'].update(value=ans_q, text_color='#9855d7')
            elif "what's up" in window['new5'].DisplayText.lower() and window['new6'].DisplayText == 'New Window':
                window['mina3'].update(source=r".\Pictures\mina80.png")
                window['new6'].update(value=ans_q, text_color='#9855d7')
            speak(ans_q)
        t = Timer(1, answer)
        t.start()
    elif 'how are you' in values['my-key3'].lower() and 'how are you' in valid_convo_starters:
        def answer():
            stmsgs = ['I am fine, how about you?!',
                      'I am fine!, lets talk about you, are you well?',
                      'I am awesome,thanks for asking', 'I am okay!, How are you?']
            ans_q = random.choice(stmsgs)
            if 'how are you' in window['new1'].DisplayText.lower() and window['new2'].DisplayText == 'New Window':
                window['mina1'].update(source=r".\Pictures\mina80.png")
                window['new2'].update(value=ans_q, text_color='#9855d7')
            elif 'how are you' in window['new3'].DisplayText.lower() and window['new4'].DisplayText == 'New Window':
                window['mina2'].update(source=r".\Pictures\mina80.png")
                window['new4'].update(value=ans_q, text_color='#9855d7')
            elif 'how are you' in window['new5'].DisplayText.lower() and window['new6'].DisplayText == 'New Window':
                window['mina3'].update(source=r".\Pictures\mina80.png")
                window['new6'].update(value=ans_q, text_color='#9855d7')
            speak(ans_q)
        t = Timer(1, answer)
        t.start()
    elif 'hey mina' in values['my-key3'].lower() and 'hey mina' in valid_convo_starters:
        def answer():
            stmsgs = ['Hey', 'Hello', 'Hi']
            ans_q = random.choice(stmsgs)
            if 'hey' in window['new1'].DisplayText.lower() and window[
                'new2'].DisplayText == 'New Window':
                window['mina1'].update(source=r".\Pictures\mina80.png")
                window['new2'].update(value=ans_q + ' ' + f'{Name.lower()}, what can i help you with?', text_color='#9855d7')
            elif 'hey' in window['new3'].DisplayText.lower() and window[
                'new4'].DisplayText == 'New Window':
                window['mina2'].update(source=r".\Pictures\mina80.png")
                window['new4'].update(value=ans_q + ' ' + f'{Name.lower()}, what can i help you with?', text_color='#9855d7')
            elif 'hey' in window['new5'].DisplayText.lower() and window[
                'new6'].DisplayText == 'New Window':
                window['mina3'].update(source=r".\Pictures\mina80.png")
                window['new6'].update(value=ans_q + ' ' + f'{Name.lower()}, what can i help you with?', text_color='#9855d7')
            speak(ans_q + ' ' + f'{Name.lower()}, what can i help you with?')
        t = Timer(1, answer)
        t.start()
    elif 'hi mina' in values['my-key3'].lower() and 'hi mina' in valid_convo_starters:
        def answer():
            stmsgs = ['Hey', 'Hello', 'Hi']
            ans_q = random.choice(stmsgs)
            if 'hi' in window['new1'].DisplayText.lower() and window[
                'new2'].DisplayText == 'New Window':
                window['mina1'].update(source=r".\Pictures\mina80.png")
                window['new2'].update(value=ans_q + ' ' + f'{Name.lower()}, what can i help you with?', text_color='#9855d7')
            elif 'hi' in window['new3'].DisplayText.lower() and window[
                'new4'].DisplayText == 'New Window':
                window['mina2'].update(source=r".\Pictures\mina80.png")
                window['new4'].update(value=ans_q + ' ' + f'{Name.lower()}, what can i help you with?', text_color='#9855d7')
            elif 'hi' in window['new5'].DisplayText.lower() and window[
                'new6'].DisplayText == 'New Window':
                window['mina3'].update(source=r".\Pictures\mina80.png")
                window['new6'].update(value=ans_q + ' ' + f'{Name.lower()}, what can i help you with?', text_color='#9855d7')
            speak(ans_q + ' ' + f'{Name.lower()}, what can i help you with?')
        t = Timer(1, answer)
        t.start()
    elif 'hello' in values['my-key3'].lower() and 'hello' in valid_convo_starters:
        def answer():
            stmsgs = ['Hey', 'Hello', 'Hi']
            ans_q = random.choice(stmsgs)
            if 'hello' in window['new1'].DisplayText.lower() and window[
                'new2'].DisplayText == 'New Window':
                window['mina1'].update(source=r".\Pictures\mina80.png")
                window['new2'].update(value=ans_q + ' ' + f'{Name.lower()}, what can i help you with?', text_color='#9855d7')
            elif 'hello' in window['new3'].DisplayText.lower() and window[
                'new4'].DisplayText == 'New Window':
                window['mina2'].update(source=r".\Pictures\mina80.png")
                window['new4'].update(value=ans_q + ' ' + f'{Name.lower()}, what can i help you with?', text_color='#9855d7')
            elif 'hello' in window['new5'].DisplayText.lower() and window[
                'new6'].DisplayText == 'New Window':
                window['mina3'].update(source=r".\Pictures\mina80.png")
                window['new6'].update(value=ans_q + ' ' + f'{Name.lower()}, what can i help you with?', text_color='#9855d7')
            speak(ans_q + ' ' + f'{Name.lower()}, what can i help you with?')
        t = Timer(1, answer)
        t.start()

    elif 'weather' in values['my-key3'].lower() and 'weather' in valid_convo_starters:
        try:
            API_KEY = "093c85c6c9b27bef3aa36af3c7743867"
            BASE_URL = "https://api.openweathermap.org/data/2.5/weather"

            requests_url = f"{BASE_URL}?appid={API_KEY}&q={location}"
            response = requests.get(requests_url)

            if response.status_code == 200:
                data = response.json()

                weather = data['weather'][0]['description']
                temperature = round(data['main']['temp'] - 273.15, 2)

                ans_q = (f"the weather in {location} is " + weather + ', with a temperature of ' +
                         str(temperature) + " degrees celsius")

                def answer():
                    if 'weather' in window['new1'].DisplayText.lower()\
                            and window['new2'].DisplayText == 'New Window':
                        window['mina1'].update(source=r".\Pictures\mina80.png")
                        window['new2'].update(value=ans_q,
                                              text_color='#9855d7')
                    elif 'weather' in window['new3'].DisplayText.lower()\
                            and window['new4'].DisplayText == 'New Window':
                        window['mina2'].update(source=r".\Pictures\mina80.png")
                        window['new4'].update(value=ans_q,
                                              text_color='#9855d7')
                    elif 'weather' in window['new5'].DisplayText.lower()\
                            and window['new6'].DisplayText == 'New Window':
                        window['mina3'].update(source=r".\Pictures\mina80.png")
                        window['new6'].update(value=ans_q,
                                              text_color='#9855d7')
                    speak(ans_q)
                t = Timer(1, answer)
                t.start()
            else:
                print("an error occurred.")
        except:
            def answer():
                if 'weather' in window['new1'].DisplayText.lower() and window[
                        'new2'].DisplayText == 'New Window':
                    window['mina1'].update(source=r".\Pictures\mina80.png")
                    window['new2'].update(value='check your internet connection',
                                          text_color='#9855d7')
                elif 'weather' in window['new3'].DisplayText.lower() and window[
                        'new4'].DisplayText == 'New Window':
                    window['mina2'].update(source=r".\Pictures\mina80.png")
                    window['new4'].update(value='check your internet connection',
                                          text_color='#9855d7')
                elif 'weather' in window['new5'].DisplayText.lower() and window[
                        'new6'].DisplayText == 'New Window':
                    window['mina3'].update(source=r".\Pictures\mina80.png")
                    window['new6'].update(value='check your internet connection',
                                          text_color='#9855d7')
                speak('check your internet connection')
            t = Timer(1, answer)
            t.start()
    elif 'maker' in values['my-key3'].lower() and 'maker' in valid_convo_starters:
        def answer():
            if 'maker' in window['new1'].DisplayText.lower() and window['new2'].DisplayText == 'New Window':
                window['mina1'].update(source=r".\Pictures\mina80.png")
                window['new2'].update(value="I was made by Richard", text_color='#9855d7')
            elif 'maker' in window['new3'].DisplayText.lower() and window['new4'].DisplayText == 'New Window':
                window['mina2'].update(source=r".\Pictures\mina80.png")
                window['new4'].update(value="I was made by Richard", text_color='#9855d7')
            elif 'maker' in window['new5'].DisplayText.lower() and window['new6'].DisplayText == 'New Window':
                window['mina3'].update(source=r".\Pictures\mina80.png")
                window['new6'].update(value="I was made by Richard", text_color='#9855d7')

        t = Timer(1, answer)
        t.start()
    elif 'creator' in values['my-key3'].lower() and 'creator' in valid_convo_starters:
        def answer():
            if 'creator' in window['new1'].DisplayText.lower() and window['new2'].DisplayText == 'New Window':
                window['mina1'].update(source=r".\Pictures\mina80.png")
                window['new2'].update(value="I was created by Richard", text_color='#9855d7')
            elif 'creator' in window['new3'].DisplayText.lower() and window['new4'].DisplayText == 'New Window':
                window['mina2'].update(source=r".\Pictures\mina80.png")
                window['new4'].update(value="I was created by Richard", text_color='#9855d7')
            elif 'creator' in window['new5'].DisplayText.lower() and window['new6'].DisplayText == 'New Window':
                window['mina3'].update(source=r".\Pictures\mina80.png")
                window['new6'].update(value="I was created by Richard", text_color='#9855d7')

        t = Timer(1, answer)
        t.start()
    elif 'joke' in values['my-key3'] and 'joke' in valid_convo_starters:

        joke = pyjokes.get_joke(language='en', category='all')

        def answer():
            if 'joke' in window['new1'].DisplayText.lower() and window['new2'].DisplayText == 'New Window':
                window['mina1'].update(source=r".\Pictures\mina80.png")
                window['new2'].update(value=joke, text_color='#9855d7')
            elif 'joke' in window['new3'].DisplayText.lower() and window['new4'].DisplayText == 'New Window':
                window['mina2'].update(source=r".\Pictures\mina80.png")
                window['new4'].update(value=joke, text_color='#9855d7')
            elif 'joke' in window['new5'].DisplayText.lower() and window['new6'].DisplayText == 'New Window':
                window['mina3'].update(source=r".\Pictures\mina80.png")
                window['new6'].update(value=joke, text_color='#9855d7')

        t = Timer(1, answer)
        t.start()
    elif "lights on" in values['my-key3'] and 'lights on' in valid_convo_starters:
        try:
            s.send("oO".encode())
            lightstate[0] = 'on'
        except:
            speak('sorry, no bluetooth device was found')

        def answer():
            if 'lights on' in window['new1'].DisplayText.lower() and window['new2'].DisplayText == 'New Window':
                window['mina1'].update(source=r".\Pictures\mina80.png")
                window['new2'].update(value='done, anything else', text_color='#9855d7')
            elif 'lights on' in window['new3'].DisplayText.lower() and window['new4'].DisplayText == 'New Window':
                window['mina2'].update(source=r".\Pictures\mina80.png")
                window['new4'].update(value='done, anything else', text_color='#9855d7')
            elif 'lights on' in window['new5'].DisplayText.lower() and window['new6'].DisplayText == 'New Window':
                window['mina3'].update(source=r".\Pictures\mina80.png")
                window['new6'].update(value='done, anything else', text_color='#9855d7')

        t = Timer(1, answer)
        t.start()
    elif "lights out" in values['my-key3'] and 'lights out' in valid_convo_starters:
        try:
            s.send("fF".encode())
            lightstate[0] = 'off'
        except:
            speak('sorry, no bluetooth device was found')

        def answer():
            if 'lights out' in window['new1'].DisplayText.lower() and window['new2'].DisplayText == 'New Window':
                window['mina1'].update(source=r".\Pictures\mina80.png")
                window['new2'].update(value='done, anything else', text_color='#9855d7')
            elif 'lights out' in window['new3'].DisplayText.lower() and window['new4'].DisplayText == 'New Window':
                window['mina2'].update(source=r".\Pictures\mina80.png")
                window['new4'].update(value='done, anything else', text_color='#9855d7')
            elif 'lights out' in window['new5'].DisplayText.lower() and window['new6'].DisplayText == 'New Window':
                window['mina3'].update(source=r".\Pictures\mina80.png")
                window['new6'].update(value='done, anything else', text_color='#9855d7')

        t = Timer(1, answer)
        t.start()
    elif "desk lamp on" in values['my-key3'] and 'desk lamp on' in valid_convo_starters:
        try:
            s.send("O".encode())
            lightstate[0] = 'on'
        except:
            speak('sorry, no bluetooth device was found')

        def answer():
            if 'desk lamp on' in window['new1'].DisplayText.lower() and window['new2'].DisplayText == 'New Window':
                window['mina1'].update(source=r".\Pictures\mina80.png")
                window['new2'].update(value='done, anything else', text_color='#9855d7')
            elif 'desk lamp on' in window['new3'].DisplayText.lower() and window['new4'].DisplayText == 'New Window':
                window['mina2'].update(source=r".\Pictures\mina80.png")
                window['new4'].update(value='done, anything else', text_color='#9855d7')
            elif 'desk lamp on' in window['new5'].DisplayText.lower() and window['new6'].DisplayText == 'New Window':
                window['mina3'].update(source=r".\Pictures\mina80.png")
                window['new6'].update(value='done, anything else', text_color='#9855d7')

        t = Timer(1, answer)
        t.start()
    elif "desk lamp off" in values['my-key3'] and 'desk lamp off' in valid_convo_starters:
        try:
            s.send("F".encode())
            lightstate[0] = 'off'
        except:
            speak('sorry, no bluetooth device was found')

        def answer():
            if 'desk lamp off' in window['new1'].DisplayText.lower() and window['new2'].DisplayText == 'New Window':
                window['mina1'].update(source=r".\Pictures\mina80.png")
                window['new2'].update(value='done, anything else', text_color='#9855d7')
            elif 'desk lamp off' in window['new3'].DisplayText.lower() and window['new4'].DisplayText == 'New Window':
                window['mina2'].update(source=r".\Pictures\mina80.png")
                window['new4'].update(value='done, anything else', text_color='#9855d7')
            elif 'desk lamp off' in window['new5'].DisplayText.lower() and window['new6'].DisplayText == 'New Window':
                window['mina3'].update(source=r".\Pictures\mina80.png")
                window['new6'].update(value='done, anything else', text_color='#9855d7')

        t = Timer(1, answer)
        t.start()
    elif "two minutes" in values['my-key3'] and 'two minutes' in valid_convo_starters:
        if lightstate[0] == 'on':
            def answer():
                if 'two minutes' in window['new1'].DisplayText.lower() and window['new2'].DisplayText == 'New Window':
                    window['mina1'].update(source=r".\Pictures\mina80.png")
                    window['new2'].update(value='Ok, will turn lights off in two minutes', text_color='#9855d7')
                elif 'two minutes' in window['new3'].DisplayText.lower() and window['new4'].DisplayText == 'New Window':
                    window['mina2'].update(source=r".\Pictures\mina80.png")
                    window['new4'].update(value='Ok, will turn lights off in two minutes', text_color='#9855d7')
                elif 'two minutes' in window['new5'].DisplayText.lower() and window['new6'].DisplayText == 'New Window':
                    window['mina3'].update(source=r".\Pictures\mina80.png")
                    window['new6'].update(value='Ok, will turn lights off in two minutes', text_color='#9855d7')

            t = Timer(1, answer)
            t.start()
            def timer():
                try:
                    s.send("fF".encode())
                    lightstate[0] = 'off'
                except:
                    speak('sorry, no bluetooth device was found')

            t = Timer(120, timer)
            t.start()
        else:
            def answer():
                if 'two minutes' in window['new1'].DisplayText.lower() and window['new2'].DisplayText == 'New Window':
                    window['mina1'].update(source=r".\Pictures\mina80.png")
                    window['new2'].update(value='Ok, will turn lights on in two minutes', text_color='#9855d7')
                elif 'two minutes' in window['new3'].DisplayText.lower() and window['new4'].DisplayText == 'New Window':
                    window['mina2'].update(source=r".\Pictures\mina80.png")
                    window['new4'].update(value='Ok, will turn lights on in two minutes', text_color='#9855d7')
                elif 'two minutes' in window['new5'].DisplayText.lower() and window['new6'].DisplayText == 'New Window':
                    window['mina3'].update(source=r".\Pictures\mina80.png")
                    window['new6'].update(value='Ok, will turn lights on in two minutes', text_color='#9855d7')

            t = Timer(1, answer)
            t.start()

            def timer():
                try:
                    s.send("oO".encode())
                    lightstate[0] = 'on'
                except:
                    speak('sorry, no bluetooth device was found')

            t = Timer(120, timer)
            t.start()
    elif 'music' in values['my-key3'] and 'music' in valid_convo_starters:
        def answer():
            music_dir2 = f'C:\\Users\\{Name}\\Music\\Playlists'
            songs = os.listdir(music_dir2)
            playlist = random.choice(songs)
            my_str = f'{playlist}'
            size = len(my_str)
            final_str = my_str[:size - 3]
            if 'music' in window['new1'].DisplayText.lower() and window['new2'].DisplayText == 'New Window':
                window['mina1'].update(source=r".\Pictures\mina80.png")
                window['new2'].update(value=f'playing {final_str}', text_color='#9855d7')
            elif 'music' in window['new3'].DisplayText.lower() and window['new4'].DisplayText == 'New Window':
                window['mina2'].update(source=r".\Pictures\mina80.png")
                window['new4'].update(value=f'playing {final_str}', text_color='#9855d7')
            elif 'music' in window['new5'].DisplayText.lower() and window['new6'].DisplayText == 'New Window':
                window['mina3'].update(source=r".\Pictures\mina80.png")
                window['new6'].update(value=f'playing {final_str}', text_color='#9855d7')
            os.startfile(os.path.join(music_dir2, playlist))

        t = Timer(1, answer)
        t.start()
    elif values['my-key3'].lower() == 'clear' and 'clear' in valid_convo_starters:
        clearchat()
        pag.press('backspace', presses=len(values['my-key3']))
    elif 'girlfriend' in values['my-key3'].lower() and 'girlfriend' in valid_convo_starters:
        def answer():
            if 'girlfriend' in window['new1'].DisplayText.lower() and window['new2'].DisplayText == 'New Window':
                window['mina1'].update(source=r".\Pictures\mina80.png")
                window['new2'].update(value='Give me some time to think about it', text_color='#9855d7')
            elif 'girlfriend' in window['new3'].DisplayText.lower() and window['new4'].DisplayText == 'New Window':
                window['mina2'].update(source=r".\Pictures\mina80.png")
                window['new4'].update(value='Give me some time to think about it', text_color='#9855d7')
            elif 'girlfriend' in window['new5'].DisplayText.lower() and window['new6'].DisplayText == 'New Window':
                window['mina3'].update(source=r".\Pictures\mina80.png")
                window['new6'].update(value='Give me some time to think about it', text_color='#9855d7')
        t = Timer(1, answer)
        t.start()
    elif 'evaluate' in values['my-key3'] and 'evaluate' in valid_convo_starters:
        query = values['my-key3'].replace('evaluate', '')
        try:
            solutions = eval(query)

            def answer():
                if 'evaluate' in window['new1'].DisplayText and \
                        window['new2'].DisplayText == 'New Window':
                    window['mina1'].update(source=r".\Pictures\mina80.png")
                    window['new2'].update(value=f'the answer is {solutions}', text_color='#9855d7')
                elif 'evaluate' in window['new3'].DisplayText and \
                        window['new4'].DisplayText == 'New Window':
                    window['mina2'].update(source=r".\Pictures\mina80.png")
                    window['new4'].update(value=f'the answer is {solutions}', text_color='#9855d7')
                elif 'evaluate' in window['new5'].DisplayText and \
                        window['new6'].DisplayText == 'New Window':
                    window['mina3'].update(source=r".\Pictures\mina80.png")
                    window['new6'].update(value=f'the answer is {solutions}', text_color='#9855d7')

            t = Timer(1, answer)
            t.start()
        except:
            pass
    elif "fine" in values['my-key3'].lower() and 'fine' in valid_convo_starters:
        def answer():
            stmsgs = ['Alright Richard', "that's good to know", "okay, what can i do your you?"]
            ans_q = random.choice(stmsgs)

            if 'fine' in window['new1'].DisplayText.lower()\
                    and window['new2'].DisplayText == 'New Window':
                window['mina1'].update(source=r".\Pictures\mina80.png")
                window['new2'].update(value=ans_q, text_color='#9855d7')
            elif 'fine' in window['new3'].DisplayText.lower() \
                    and window['new4'].DisplayText == 'New Window':
                window['mina2'].update(source=r".\Pictures\mina80.png")
                window['new4'].update(value=ans_q, text_color='#9855d7')
            elif 'fine' in window['new5'].DisplayText.lower()\
                    and window['new6'].DisplayText == 'New Window':
                window['mina3'].update(source=r".\Pictures\mina80.png")
                window['new6'].update(value=ans_q, text_color='#9855d7')
        t = Timer(1, answer)
        t.start()
    elif "awesome" in values['my-key3'].lower() and 'awesome' in valid_convo_starters:
        def answer():
            stmsgs = ['Alright Richard', "that's good to know", "okay, what can i do your you?"]
            ans_q = random.choice(stmsgs)

            if 'awesome' in window['new1'].DisplayText.lower()\
                    and window['new2'].DisplayText == 'New Window':
                window['mina1'].update(source=r".\Pictures\mina80.png")
                window['new2'].update(value=ans_q, text_color='#9855d7')
            elif 'awesome' in window['new3'].DisplayText.lower() \
                    and window['new4'].DisplayText == 'New Window':
                window['mina2'].update(source=r".\Pictures\mina80.png")
                window['new4'].update(value=ans_q, text_color='#9855d7')
            elif 'awesome' in window['new5'].DisplayText.lower()\
                    and window['new6'].DisplayText == 'New Win#9855ddow':
                window['mina3'].update(source=r".\Pictures\mina80.png")
                window['new6'].update(value=ans_q, text_color='#9855d7')
        t = Timer(1, answer)
        t.start()
    elif "great" in values['my-key3'].lower() and 'great' in valid_convo_starters:
        def answer():
            stmsgs = ['Alright Richard', "that's good to know", "okay, what can i do your you?"]
            ans_q = random.choice(stmsgs)

            if 'great' in window['new1'].DisplayText.lower()\
                    and window['new2'].DisplayText == 'New Window':
                window['mina1'].update(source=r".\Pictures\mina80.png")
                window['new2'].update(value=ans_q, text_color='#9855d7')
            elif 'great' in window['new3'].DisplayText.lower() \
                    and window['new4'].DisplayText == 'New Window':
                window['mina2'].update(source=r".\Pictures\mina80.png")
                window['new4'].update(value=ans_q, text_color='#9855d7')
            elif 'great' in window['new5'].DisplayText.lower()\
                    and window['new6'].DisplayText == 'New Window':
                window['mina3'].update(source=r".\Pictures\mina80.png")
                window['new6'].update(value=ans_q, text_color='#9855d7')
        t = Timer(1, answer)
        t.start()
    elif "alright" in values['my-key3'].lower() and 'alright' in valid_convo_starters:
        def answer():
            stmsgs = ['Alright Richard', "that's good to know", "okay, what can i do your you?"]
            ans_q = random.choice(stmsgs)

            if 'alright' in window['new1'].DisplayText.lower()\
                    and window['new2'].DisplayText == 'New Window':
                window['mina1'].update(source=r".\Pictures\mina80.png")
                window['new2'].update(value=ans_q, text_color='#9855d7')
            elif 'alright' in window['new3'].DisplayText.lower() \
                    and window['new4'].DisplayText == 'New Window':
                window['mina2'].update(source=r".\Pictures\mina80.png")
                window['new4'].update(value=ans_q, text_color='#9855d7')
            elif 'alright' in window['new5'].DisplayText.lower()\
                    and window['new6'].DisplayText == 'New Window':
                window['mina3'].update(source=r".\Pictures\mina80.png")
                window['new6'].update(value=ans_q, text_color='#9855d7')
        t = Timer(1, answer)
        t.start()
    elif "good" in values['my-key3'].lower() and 'good' in valid_convo_starters:
        def answer():
            stmsgs = ['Alright Richard', "that's good to know", "okay, what can i do your you?"]
            ans_q = random.choice(stmsgs)

            if 'good' in window['new1'].DisplayText.lower()\
                    and window['new2'].DisplayText == 'New Window':
                window['mina1'].update(source=r".\Pictures\mina80.png")
                window['new2'].update(value=ans_q, text_color='#9855d7')
            elif 'good' in window['new3'].DisplayText.lower() \
                    and window['new4'].DisplayText == 'New Window':
                window['mina2'].update(source=r".\Pictures\mina80.png")
                window['new4'].update(value=ans_q, text_color='#9855d7')
            elif 'good' in window['new5'].DisplayText.lower()\
                    and window['new6'].DisplayText == 'New Window':
                window['mina3'].update(source=r".\Pictures\mina80.png")
                window['new6'].update(value=ans_q, text_color='#9855d7')
        t = Timer(1, answer)
        t.start()
    else:
        # if the input is not part of the list of valid convo starters
        # it reroutes the input to the openai response generator
        def answer():
            if window['new1'].DisplayText.lower() not in valid_convo_starters \
                    and window['new2'].DisplayText == 'New Window':
                try:
                    resp = generate_response(window['new1'].DisplayText)
                    window['mina1'].update(source=r".\Pictures\mina80.png")
                    window['new2'].update(value=resp, text_color='#9855d7')
                    speak(resp)
                except Exception:
                    resp = "check your internet connection"
                    window['mina1'].update(source=r".\Pictures\mina80.png")
                    window['new2'].update(value=resp, text_color='#9855d7')
                    speak(resp)

            elif window['new3'].DisplayText.lower() not in valid_convo_starters \
                    and window['new4'].DisplayText == 'New Window':
                try:
                    resp = generate_response(window['new3'].DisplayText)
                    window['mina2'].update(source=r".\Pictures\mina80.png")
                    window['new4'].update(value=resp, text_color='#9855d7')
                    speak(resp)
                except Exception:
                    resp = "check your internet connection"
                    window['mina2'].update(source=r".\Pictures\mina80.png")
                    window['new4'].update(value=resp, text_color='#9855d7')
                    speak(resp)
            elif window['new5'].DisplayText.lower() not in valid_convo_starters \
                    and window['new6'].DisplayText == 'New Window':
                try:
                    resp = generate_response(window['new5'].DisplayText)
                    window['mina3'].update(source=r".\Pictures\mina80.png")
                    window['new6'].update(value=resp, text_color='#9855d7')
                    speak(resp)
                except Exception:
                    resp = "check your internet connection"
                    window['mina3'].update(source=r".\Pictures\mina80.png")
                    window['new6'].update(value=resp, text_color='#9855d7')
                    speak(resp)

        def answer2():
            if window['new1'].DisplayText.lower() not in valid_convo_starters \
                    and window['new2'].DisplayText == 'New Window':
                window['mina1'].update(source=r".\Pictures\mina80.png")
                window['new2'].update(value='check your internet connection', text_color='#9855d7')
            elif window['new3'].DisplayText.lower() not in valid_convo_starters \
                    and window['new4'].DisplayText == 'New Window':
                window['mina2'].update(source=r".\Pictures\mina80.png")
                window['new4'].update(value='check your internet connection', text_color='#9855d7')
            elif window['new5'].DisplayText.lower() not in valid_convo_starters \
                    and window['new6'].DisplayText == 'New Window':
                window['mina3'].update(source=r".\Pictures\mina80.png")
                window['new6'].update(value='check your internet connection', text_color='#9855d7')
            speak('check your internet connection')
        if check_internet():
            # runs function one if there is an internet connection
            t = Timer(1/2, answer)
            t.start()
        else:
            # runs function two if there is no internet connection
            t = Timer(1/2, answer2)
            t.start()


def updatetext():
    hour = int(datetime.datetime.now().hour)
    if 0 <= hour < 12:
        window['text1'].update("--Good Morning!. I'm Mina. A virtual assistant\n for windows")
    elif 12 <= hour < 18:
        window['text1'].update("--Good Afternoon!. I'm Mina. A virtual assistant\n for windows")
    else:
        window['text1'].update("--Good Evening!. I'm Mina. A virtual assistant\n for windows")


updatetext()


def checkProcessRunning(processName):
    # Checking if there is any running process that contains the given name processName.
    # Iterate over the all the running process
    for proc in psutil.process_iter():
        try:
            # Check if process name contains the given name string.
            if processName.lower() in proc.name().lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return False


if __name__ == '__main__':
    try:
        os.startfile(r'.\SpeechRecognizer.exe')
        while True:
            event, values = window.read()
            window.bind("<Return>", "-ENTER-")
            window.bind("<Escape>", "-ESC-")
            if event == 'btnkey':
                if checkProcessRunning('SpeechRecognizer'):
                    subprocess.call("TASKKILL /F /IM SpeechRecognizer.exe", shell=True)
                    window['btnkey'].update(image_filename=r".\Pictures\recog-mic.png")
                else:
                    os.startfile(r'.\SpeechRecognizer.exe')
                    window['btnkey'].update(image_filename=r".\Pictures\recog-mic2.png")
            elif event == 'mode':
                subprocess.call("TASKKILL /F /IM SpeechRecognizer.exe", shell=True)
                os.startfile(r'.\SpeechRecognizer2.exe')
            elif event == 'unrgstr':
                os.remove(r'.\encodename.pickle')
                os.remove(r'.\encodecity.pickle')
                os.remove(r'.\encodecountry.pickle')
                os.startfile(r".\Profile Registry.exe")
            elif 'open YouTube' in values['my-key'] and event == "-ENTER-":
                webbrowser.open('https://www.youtube.com/')
            elif event == 'scan':
                open_scanner()
            elif event == 'con':
                try:
                    if os.path.exists(r'.\bluetoothaddress.pickle'):
                        address = pickle.load(open('bluetoothaddress.pickle', 'rb'))
                        addr = address  # Device Address
                        port = 1  # RFCOMM port
                        passkey = "5067"  # passkey of the device you want to connect

                        # Now, connect in the same way as always with PyBlueZ
                        s = bluetooth.BluetoothSocket(bluetooth.RFCOMM)
                        s.connect((addr, port))
                        speak('Bluetooth device connected successfully')
                except:
                    speak('Bluetooth Connection failed')
            elif 'open browser' in values['my-key'] and event == "-ENTER-" or \
                    values['my-key'] == 'open browser' and event == "OK":
                os.startfile("C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe")
            elif values['my-key'] == 'VS code' and event == "-ENTER-" or values['my-key'] == 'VS code' and event == "OK":
                try:
                    os.startfile(f"C:\\Users\\{Name}\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe")
                except:
                    pass
            elif values['my-key2'] and event == "SEARCH" or values['my-key2'] and event == "-ENTER-":
                webbrowser.open('https://www.google.com/search?q=' + values['my-key2'])
            elif values['my-key'] == 'VLC' and event == "-ENTER-" or values['my-key'] == 'VLC' and event == "OK":
                try:
                    os.startfile("C:\\Program Files\\VideoLAN\\VLC\\vlc.exe")
                except:
                    pass
            elif values['my-key'] == 'music' and event == "-ENTER-" or values['my-key'] == 'music' and event == "OK":
                os.startfile(f"C:\\Users\\{Name}\\Music")
            elif values['my-key'] == 'videos from disk d' and event == "-ENTER-" or \
                    values['my-key'] == 'videos from disk d' and event == "OK":
                os.startfile("D:\\Videos\\Movies")

            elif values['my-key'] == 'Videos' and event == "-ENTER-" or values['my-key'] == 'Videos' and event == "OK":
                os.startfile(f"C:\\Users\\{Name}\\Videos")
            elif values['my-key'] == 'winamp' and event == "-ENTER-" or values['my-key'] == 'winamp' and event == "OK":
                os.startfile("C:\\Program Files (x86)\\Winamp\\winamp.exe")
            elif values['my-key'] == 'volume down' and event == "OK":
                pag.press('volumedown')
            elif values['my-key'] == 'volume mute' and event == "OK":
                pag.press('volumemute')
            elif values['my-key'] == 'play music' and event == "-ENTER-":
                play_music()
            elif values['my-key'] and event == "-ESC-":
                client = wolframalpha.Client('HAGRHY-KVWGAWWA8R')
                value = values['my-key']
                length = len(value)
                new_value = value[:length - 1]
                res = client.query(new_value)
                try:
                    answer = next(res.results).text
                    print(answer)
                    speak(answer)
                except StopIteration:
                    print("No results")
            elif event == "Commands":
                window6 = sG.popup_scrolled('''Voice commands:
                                             'open Google',
                                             'open YouTube',
                                             'Music','Videos',
                                             'Videos from disk D',
                                             'check the time',
                                             'open browser','VLC',
                                              'Krita','VS Code' ''',
                                            '''Text commands:
                                               'open Google',
                                               'open YouTube',
                                               'Music','Videos',
                                               'Videos D',
                                               'check the time',
                                               'open browser',
                                               'Krita','VS Code'
                                               'VLC'
                                            between your words i.e 'how+to+make+bread' ''')
            elif event == "-ENTER-" and values['my-key3'] != '':
                if window['new6'].DisplayText == 'New Window':
                    if 'New Window' in window['new1'].DisplayText:
                        window['myprofile'].update(source=r".\Pictures\R.png")
                        window['new1'].update(value=values['my-key3'], text_color='#00baff')
                    elif window['new1'].DisplayText != 'New Window' and window['new3'].DisplayText == 'New Window':
                        window['myprofile2'].update(source=r".\Pictures\R.png")
                        window['new3'].update(value=values['my-key3'], text_color='#00baff')
                    elif window['new1'].DisplayText != 'New Window' and window['new3'] != 'New Window':
                        window['myprofile3'].update(source=r".\Pictures\R.png")
                        window['new5'].update(value=values['my-key3'], text_color='#00baff')
                    pag.press('backspace', presses=len(values['my-key3']))
                    chatbot()
                elif window['new6'].DisplayText != 'New Window':
                    clearchat()
                    window['myprofile'].update(source=r".\Pictures\R.png")
                    window['new1'].update(value=values['my-key3'], text_color='#00baff')
                    pag.press('backspace', presses=len(values['my-key3']))
                    chatbot()

            elif event == "EXIT":
                subprocess.call("TASKKILL /F /IM SpeechRecognizer.exe", shell=True)
                subprocess.call("TASKKILL /F /IM SpeechRecognizer2.exe", shell=True)
                break
            elif event == sG.WIN_CLOSED:
                subprocess.call("TASKKILL /F /IM SpeechRecognizer.exe", shell=True)
                subprocess.call("TASKKILL /F /IM SpeechRecognizer2.exe", shell=True)
                break
        window.close()
    except:
        pass
