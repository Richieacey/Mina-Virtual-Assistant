import os
import pickle
import PySimpleGUI as sg
from playsound import playsound


def set_location(city, country):
    # user'sCity
    file = open(r'./encodecity.pickle', 'wb')
    # dump encoding in pickle file
    pickle.dump(city, file)
    file.close()
    # user'sCountry
    file = open(r'./encodecountry.pickle', 'wb')
    # dump encoding in pickle file
    pickle.dump(country, file)
    file.close()


def set_name(name):
    file = open(r'./encodename.pickle', 'wb')
    # dump encoding in pickle file
    pickle.dump(name, file)
    file.close()



layout = [
          [sg.Text('First Name:', font=('Myanmar Text', 16), background_color='#0bec9a', pad=((30, 0), (20, 0)))],
          [sg.InputText(key='-IN4-', size=(10, 18), font=('Arial', 18), background_color="#00baff",
                        enable_events=True, pad=((30, 0), (20, 0)))],
            [sg.Text('City:',font=('Myanmar Text', 16), background_color='#0bec9a', pad=((30, 0), (20, 0)))],
          [sg.InputText(key='-IN1-',  size=(10, 18),font=('Arial', 18), background_color="#00baff",
                        enable_events=True, pad=((30, 0), (20, 0)))],
            [sg.Text('Country:', font=('Myanmar Text', 16), background_color='#0bec9a', pad=((30, 0), (20, 0)))],
            [sg.InputText(key='-IN2-', size=(10, 18), font=('Arial', 18), background_color="#00baff",
                          enable_events=True, pad=((30, 0), (20, 0)))],
          [sg.Button("Register", pad=((35, 20), (0, 15))), sg.Button("Exit", button_color='red', pad=((20, 0), (0, 15)))]]

window = sg.Window("MINA-Profile Registry", layout, resizable=True, background_color="#0bec9a",
                   icon=r".\Pictures\lock.ico", titlebar_background_color="#00baff",
                   titlebar_font="", titlebar_text_color="white", use_custom_titlebar=True,
                   titlebar_icon=r".\Pictures\lock.png",
                   finalize=True, modal=False)
if __name__ == '__main__':
    try:
        while True:
            event, values = window.read()
            if event in (sg.WIN_CLOSED, 'Exit'):
                break
            elif event == 'Register':
                set_location(values['-IN1-'], values['-IN2-'])
                set_name(values['-IN4-'])
                playsound(r'./Unlock Item  Sound Effect HD.mp3')
    except:
        pass
