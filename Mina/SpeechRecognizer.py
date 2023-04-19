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
import wikipedia
from vosk import Model, KaldiRecognizer
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
import bluetooth
import pickle
import openai
import numpy as np


OPENAI_API_KEY = "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"
openai.api_key = OPENAI_API_KEY

# initialize text to speech engine
engine = py.init('sapi5')
voices = engine.getProperty('voices')
# sets the voice to female
engine.setProperty('voice', voices[1].id)

WAKE = "mina"

# vosk speech recognition model which has offline recognition capability
model = Model(r".\vosk-model-small-en-us-0.15\vosk-model-small-en-us-0.15")

# sets the volume of the text to speech engine(from 0 to 1)
engine.setProperty("volume", 1)

# set the speed of the text to speech engine
rate = engine.getProperty('rate')
engine.setProperty('rate', 180)

# Get the computer name, and this is the name the assistant refers to you by
Name = win32api.GetComputerName()
City = 'New york'
Country = "America"


# create a function for the text to speech engine
def speak(audio1):
    engine.say(audio1)
    engine.runAndWait()


if os.path.exists('encodename.pickle'):
    Name = pickle.load(open('encodename.pickle', 'rb'))

if os.path.exists('encodecity.pickle'):
    City = pickle.load(open('encodecity.pickle', 'rb'))

if os.path.exists('encodecountry.pickle'):
    Country = pickle.load(open('encodecountry.pickle', 'rb'))


# tries to connect a bluetooth device for home automation
try:
    # searches for the pickle file that was made with the MINA GUI script
    if os.path.exists(r'.\bluetoothaddress.pickle'):
        address = pickle.load(open('bluetoothaddress.pickle', 'rb'))
        addr = address     # Device Address
        port = 1         # RFCOMM port

        # Now, connect in the same way as always with PyBlueZ
        s = bluetooth.BluetoothSocket(bluetooth.RFCOMM)
        s.connect((addr, port))
        speak('Bluetooth device connected successfully')
except Exception:
    speak('Bluetooth Connection failed')


# Openai api function for generating response
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


# check the time
def asked_time():
    now = datetime.datetime.now()
    dts = now.strftime('%I:%M')
    current_time = dts
    hour = int(datetime.datetime.now().hour)
    if 0 <= hour < 12:
        speak(f"the time is {current_time + ' AM'}")
    else:
        speak(f"the time is {current_time + ' PM'}")


# play music function
def play_music():
    # access the playlists folder
    music_dir2 = f'C:\\Users\\{Name}\\Music\\Playlists'
    # creates a list of songs(playlists)
    songs = os.listdir(music_dir2)
    playlist = random.choice(songs)
    my_str = f'{playlist}'
    size = len(my_str)
    # removes the extension name
    final_str = my_str[:size - 3]
    # speak the playlist that was randomly selected
    speak(f'playing {final_str}')
    # use the os module to startfile with the default music player
    os.startfile(os.path.join(music_dir2, playlist))


# weather update function
def weather():
    try:
        API_KEY = "093c85c6c9b27bef3aa36af3c7743867"
        BASE_URL = "https://api.openweathermap.org/data/2.5/weather"

        requests_url = f"{BASE_URL}?appid={API_KEY}&q={City}"
        response = requests.get(requests_url)

        if response.status_code == 200:
            data = response.json()
            weather = data['weather'][0]['description']
            temperature = round(data['main']['temp'] - 273.15, 2)

            speak("the weather in Benin City is " + weather + ', with a temperature of '
                  + str(temperature) + " degree celsius")
            print(data)
        else:
            print("an error occurred.")
    except Exception:
        speak("check your internet connection")


# news from the international community
def NewsFromBBC():
    try:
        # BBC news api
        # following query parameters are used
        # source, sortBy and apiKey
        query_params = {
            "source": "bbc-news",
            "country": f"{Country}",
            "sortBy": "top",
            "apiKey": "6303a0d3dafb4d72a4eff925047c8e38"
        }
        main_url = " https://newsapi.org/v1/articles"

        # fetching data in json format
        res = requests.get(main_url, params=query_params)
        open_bbc_page = res.json()

        # getting all articles in a string article
        article = open_bbc_page["articles"]

        # empty list which will
        # contain all trending news
        results = []
        urls = []

        for ar in article:
            results.append(ar["title"])
            results.append(ar["description"])
        for url in article:
            urls.append(url["url"])

        for i in range(len(results)):
            # printing all trending news
            print(i + 1, results[i])

        for i in range(len(urls)):
            print(i + 1, urls[i])
        speak(results)
    except Exception:
        speak("check your internet connection")


# main speech recognizer function
def speech_recognizer():
    # intialize KaldiRecognizer with the vosk model
    recognizer2 = KaldiRecognizer(model, 16000)

    # initialize microphone with pyaudio module
    mic = pyaudio.PyAudio()
    stream = mic.open(format=pyaudio.paInt16, channels=1, rate=16000, input=True, frames_per_buffer=8192)
    # start mic stream
    stream.start_stream()
    # first instance of the recognizer "while loop" to get the wake word
    while True:
        # read data from the microphone
        data = stream.read(4096)
        if recognizer2.AcceptWaveform(data):
            # speech recognizer result
            text_msg1 = recognizer2.Result()
            # if the wake word in result
            if WAKE in text_msg1:
                speak("I am here,what can i help you with?")
                # second instance of the recognizer "while loop" runs only when the wake word is correct
                while True:
                    data = stream.read(4096)
                    if recognizer2.AcceptWaveform(data):
                        text_msg1 = recognizer2.Result()
                        # vosk recognizer returns the result in a weird format, so i have to strip all the
                        # uncessary characters in the string
                        text_msg = text_msg1[14:-3]
                        if "what's up" in text_msg:
                            # list of possible responses
                            stmsgs = ['as of now i\'m only helping you with whatever you need me for',
                                      'I am fine!, lets talk about you, are you fine?', 'oh, nothing',
                                      'i am okay ! How are you?']
                            # randomly selects from the list
                            ans_q = random.choice(stmsgs)
                            # speak selection
                            speak(ans_q)
                        elif 'google' in text_msg:  # opens google with the default browser
                            try:
                                webbrowser.open('https://www.google.com/')
                            except Exception:
                                speak("check your internet connection")
                        elif "wiki" in text_msg:  # check wikipedia for information
                            try:
                                speak("searching details...")
                                text_msg.replace("wiki", "")
                                results = wikipedia.summary(text_msg, sentences=10)
                                f = open('new.txt', 'w+')
                                f.write(results)
                                f.close()
                                speak(results)
                                # a text document to store the information
                                os.startfile(r".\new.txt")
                            except Exception:
                                speak(f"could not find{text_msg}")
                        elif WAKE in text_msg and "weather" in text_msg:  # calls the weather inquiry function
                            speak("looking for the weather...")
                            weather()
                        elif "check the news" in text_msg:
                            speak(f"alright {Name}, checking for news from around the world")
                            NewsFromBBC()
                        elif text_msg == "open you tube":
                            try:
                                webbrowser.open("https://www.youtube.com")
                            except Exception:
                                speak("check your internet connection")
                        elif "pandora" in text_msg:
                            pag.hotkey('ctrl', 'alt', 'p')
                            speak(f'alright {Name},locking work station now')
                        elif WAKE in text_msg and "how are you" in text_msg:  # pre-programmed small talk
                            stmsgs = ['I am fine, how about you?!',
                                      'I am fine!, lets talk about you, are you well?',
                                      'I am awesome,thanks for asking', 'I am okay!, How are you?']
                            ans_q = random.choice(stmsgs)
                            speak(ans_q)
                        elif "your maker" in text_msg:
                            speak("i was made by Richard")
                        elif "your creator" in text_msg:
                            speak("i was created by Richard")
                        elif text_msg == "mina":
                            speak(f"Yes {Name}.how may I help you?")
                        # home automation commands
                        # uses the bluetooth module to send strings to a HC-O5 bluetooth module connected to arduino
                        elif "all lights on" in text_msg:
                            try:
                                s.send("oO".encode())
                            except Exception:
                                speak('sorry, no bluetooth device was found')
                        elif "lights out" in text_msg:
                            try:
                                s.send("fF".encode())
                            except Exception:
                                speak('sorry, no bluetooth device was found')
                        elif "desk lamp on" in text_msg:
                            try:
                                s.send("O".encode())
                            except Exception:
                                speak('sorry, no bluetooth device was found')
                        elif "desk lamp off" in text_msg:
                            try:
                                s.send("F".encode())
                            except Exception:
                                speak('sorry, no bluetooth device was found')
                        elif text_msg == "lights on":
                            try:
                                s.send("o".encode())
                            except Exception:
                                speak('sorry, no bluetooth device was found')
                        elif text_msg == "lights off":
                            try:
                                s.send("f".encode())
                            except Exception:
                                speak('sorry, no bluetooth device was found')
                        elif text_msg == "go to sleep":  # breaks the while second loop, but keeps the first active
                            # can be woken with the wake word, and the second while loop starts again
                            speak('going to sleep, wake me when you need me')
                            break
                        elif "bored" in text_msg:  # if the user says they are bored, tell a joke or play music
                            speak("would you like to hear a joke?")
                            # this while loop allows for multiple inquiries from the assistant
                            while True:
                                data1 = stream.read(4096)
                                if recognizer2.AcceptWaveform(data1):
                                    text_msg1 = recognizer2.Result()
                                    command = text_msg1[14:-3]
                                    if "yes" in command:
                                        joke = pyjokes.get_joke(language='en', category='neutral')
                                        print(joke)
                                        speak(joke)
                                    elif "no" in command:
                                        speak('how about some music?')
                                        while True:
                                            data = stream.read(4096)
                                            if recognizer2.AcceptWaveform(data):
                                                command2 = recognizer2.Result()
                                                if "yes" in command2:
                                                    play_music()
                                                elif "no" in command2:
                                                    speak("Okay, let me know if there's anything"
                                                          " else i can help you with")
                                                break
                                    break
                        elif "tell me a joke" in text_msg:  # explicit request for jokes, replies using
                            # pyjokes module
                            joke = pyjokes.get_joke(language='en', category='neutral')
                            print(joke)
                            speak(joke)
                        elif text_msg == "stop music":  # runs a taskkill call on the possible music players the user
                            # may be using
                            try:
                                subprocess.call("TASKKILL /F /IM wmplayer.exe", shell=True)
                            except Exception:
                                pass
                            try:
                                subprocess.call("TASKKILL /F /IM winamp.exe", shell=True)
                            except Exception:
                                pass
                        elif "screenshot" in text_msg:  # takes a screenshot with the pyautogui module
                            scrs = pag.screenshot()
                            # the image is saved to desktop with the name "screenshot.png"
                            # this has to be renamed by the user, so that the next screenshot will not override
                            # the previous
                            scrs.save(r"C:\Users\RICHARD\Desktop\screenshot.png")
                            os.startfile(r"C:\Users\RICHARD\Desktop\screenshot.png")
                        elif "i'm fine" in text_msg:
                            speak("oh that's nice, anyways what can i help you with today?")
                        elif text_msg == "hey mina":
                            stmsgs = ['hi how can i help you today', 'hello there, what can i do for you',
                                      'hey, what can i help you with today',
                                      'hello, what do you want me to do today']
                            ans_q = random.choice(stmsgs)
                            speak(ans_q)
                        elif "girlfriend" in text_msg:  # most popular question users ask assistants
                            speak("give me some days to think about it")
                        elif "minimize" in text_msg:  # this command returns us to the Desktop
                            pag.hotkey('winleft', 'd')
                        elif text_msg == 'music folder':  # opens the music folder by inserting the computer's name
                            # in the directory
                            os.startfile(f"C:\\Users\\{Name}\\Music")
                        elif 'recycle' in text_msg:  # emptying the recycle bin with the winshell function
                            speak("emptying recycle bin")
                            try:
                                winshell.recycle_bin().empty(confirm=False, show_progress=True, sound=True)
                                speak("done,anything else?")
                            except Exception:
                                speak(f"recycle bin is currently empty {Name}")
                        elif "videos" in text_msg: # opens the videos folder by inserting the computer's name
                            # in the directory
                            os.startfile(f"C:\\Users\\{Name}\\Videos")
                        elif 'check the time' in text_msg:  # this command runs the asked time function declared above
                            asked_time()
                        elif 'open vs code' in text_msg:  # a little something for the developers
                            # this opens vs code provided the folder location is not changed
                            os.startfile(f"C:\\Users\\{Name}\\"
                                         f"AppData\\Local\\Programs\\Microsoft VS "
                                         "Code\\Code.exe")
                        elif 'open browser' in text_msg:  # opens the default browser
                            os.startfile("C:\\Program Files (x86)\\"
                                         "Microsoft\\Edge\\Application\\msedge.exe")
                        elif 'play music' in text_msg: # play music function
                            play_music()
                        elif text_msg == 'volume down':  # volume control with pyautogui module
                            # tested it out, 25 presses reduces the volume to 50%
                            pag.press("volumedown", presses=25)
                        elif text_msg == 'volume up':  # this just keeps increasing it until the 80 presses are done
                            pag.press('volumeup', presses=80)
                        elif text_msg == 'volume to twenty percent': # this will only work if the volume was at 100%
                            # to begin with
                            pag.press("volumedown", presses=40)
                        elif text_msg == 'volume to seventy percent': # the same applies here
                            pag.press("volumedown", presses=15)
                        elif text_msg == "pause":
                            pag.press('playpause', presses=1)
                        elif text_msg == "play":  # uses the pyautogui module to press a virtual playpause button
                            # on the keyboard
                            pag.press("playpause")
                        elif text_msg == "next": # presses the nexttrack virtual button
                            pag.press("nexttrack")
                        elif text_msg == "previous": # previous tract button
                            pag.press("prevtrack")
                        elif "playlist number two" in text_msg:  # these will work if you the position of each
                            # playlist in the list, I did this so others can use it instead of doing something like
                            # "Michael jackson in text_msg"
                            try:
                                music_dir2 = fr'C:\Users\{Name}\Music\Playlists'
                                songs = os.listdir(music_dir2)
                                my_str = f'{songs[1]}'
                                size = len(my_str)
                                final_str = my_str[:size - 3]
                                speak(f'playing {final_str}')
                                os.startfile(os.path.join(music_dir2, songs[1]))
                            except Exception:
                                speak("There's no such playlist in your directory,"
                                      " you can create it if like.")
                        elif "playlist number one" in text_msg:
                            try:
                                music_dir2 = fr'C:\Users\{Name}\Music\Playlists'
                                songs = os.listdir(music_dir2)
                                my_str = f'{songs[0]}'
                                size = len(my_str)
                                final_str = my_str[:size - 3]
                                speak(f'playing {final_str}')
                                os.startfile(os.path.join(music_dir2, songs[0]))
                            except Exception:
                                speak("There's no such playlist in your directory,"
                                      " you can create it if like.")
                        elif "playlist number three" in text_msg:
                            try:
                                music_dir2 = fr'C:\Users\{Name}\Music\Playlists'
                                songs = os.listdir(music_dir2)
                                my_str = f'{songs[2]}'
                                size = len(my_str)
                                final_str = my_str[:size - 3]
                                speak(f'playing {final_str}')
                                os.startfile(os.path.join(music_dir2, songs[2]))
                            except Exception:
                                speak("There's no such playlist in your directory,"
                                      " you can create it if like.")
                        elif "playlist number four" in text_msg:
                            try:
                                music_dir2 = fr'C:\Users\{Name}\Music\Playlists'
                                songs = os.listdir(music_dir2)
                                my_str = f'{songs[3]}'
                                size = len(my_str)
                                final_str = my_str[:size - 3]
                                speak(f'playing {final_str}')
                                os.startfile(os.path.join(music_dir2, songs[3]))
                            except Exception:
                                speak("There's no such playlist in your directory,"
                                      " you can create it if like.")
                        elif "playlist number five" in text_msg:
                            try:
                                music_dir2 = fr'C:\Users\{Name}\Music\Playlists'
                                songs = os.listdir(music_dir2)
                                my_str = f'{songs[4]}'
                                size = len(my_str)
                                final_str = my_str[:size - 3]
                                speak(f'playing {final_str}')
                                os.startfile(os.path.join(music_dir2, songs[4]))
                            except Exception:
                                speak("There's no such playlist in your directory,"
                                      " you can create it if like.")
                        elif "playlist number six" in text_msg:
                            try:
                                music_dir2 = fr'C:\Users\{Name}\Music\Playlists'
                                songs = os.listdir(music_dir2)
                                my_str = f'{songs[5]}'
                                size = len(my_str)
                                final_str = my_str[:size - 3]
                                speak(f'playing {final_str}')
                                os.startfile(os.path.join(music_dir2, songs[5]))
                            except Exception:
                                speak("There's no such playlist in your directory,"
                                      " you can create it if like.")
                        elif "playlist number seven" in text_msg:
                            try:
                                music_dir2 = fr'C:\Users\{Name}\Music\Playlists'
                                songs = os.listdir(music_dir2)
                                my_str = f'{songs[6]}'
                                size = len(my_str)
                                final_str = my_str[:size - 3]
                                speak(f'playing {final_str}')
                                os.startfile(os.path.join(music_dir2, songs[6]))
                            except Exception:
                                speak("There's no such playlist in your directory, "
                                      "you can create it if like.")
                        elif "playlist number eight" in text_msg:
                            try:
                                music_dir2 = fr'C:\Users\{Name}\Music\Playlists'
                                songs = os.listdir(music_dir2)
                                my_str = f'{songs[7]}'
                                size = len(my_str)
                                final_str = my_str[:size - 3]
                                speak(f'playing {final_str}')
                                os.startfile(os.path.join(music_dir2, songs[7]))
                            except Exception:
                                speak("There's no such playlist in your directory,"
                                      " you can create it if like.")
                        elif "number nine" in text_msg:
                            try:
                                music_dir2 = fr'C:\Users\{Name}\Music\Playlists'
                                songs = os.listdir(music_dir2)
                                my_str = f'{songs[8]}'
                                size = len(my_str)
                                final_str = my_str[:size - 3]
                                speak(f'playing {final_str}')
                                os.startfile(os.path.join(music_dir2, songs[8]))
                            except Exception:
                                speak("There's no such playlist in your directory, "
                                      "you can create it if like.")
                        elif "playlist number ten" in text_msg:
                            try:
                                music_dir2 = fr'C:\Users\{Name}\Music\Playlists'
                                songs = os.listdir(music_dir2)
                                my_str = f'{songs[9]}'
                                size = len(my_str)
                                final_str = my_str[:size - 3]
                                speak(f'playing {final_str}')
                                os.startfile(os.path.join(music_dir2, songs[9]))
                            except Exception:
                                speak("There's no such playlist in your directory,"
                                      " you can create it if like.")
                        elif 'mina system down' in text_msg:  # this runs a subprocess call that shuts down your
                            # computer
                            speak('shutting down....')
                            subprocess.call('shutdown /p /f')
                        elif "question mode" in text_msg:
                            while True:
                                data1 = stream.read(4096)
                                if recognizer2.AcceptWaveform(data1):
                                    text_msg1 = recognizer2.Result()
                                    command = text_msg1[14:-3]
                                    print(command)
                                    if "command mode" in command:
                                        break
                                    else:
                                        try:
                                            response = generate_response(command)
                                            speak(response)
                                        except Exception:
                                            speak('internet not connected')
                                            speak('switching back to command mode')
                                            break


# checks for an internet connection
def check_internet(host="https://www.google.com/"):
    try:
        urllib.request.urlopen(host)
        return True
    except Exception:
        return False


if check_internet():
    speak('connection successful')
else:
    speak('failed to connect to a network, you may not be able to access all of my functions')

if __name__ == '__main__':
    speech_recognizer()