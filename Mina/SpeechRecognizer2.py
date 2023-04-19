import pyttsx3 as py
from vosk import Model, KaldiRecognizer
import openai
import os
import subprocess
import pyaudio


OPENAI_API_KEY = "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"
openai.api_key = OPENAI_API_KEY
# vosk speech recognition model which has offline recognition capability
model = Model(r".\vosk-model-small-en-us-0.15\vosk-model-small-en-us-0.15")

# initialize text to speech engine
engine = py.init('sapi5')
voices = engine.getProperty('voices')
# sets the voice to female
engine.setProperty('voice', voices[1].id)

# sets the volume of the text to speech engine(from 0 to 1)
engine.setProperty("volume", 1)

# set the speed of the text to speech engine
rate = engine.getProperty('rate')
engine.setProperty('rate', 180)

# Get the computer name, and this is the name the assistant refers to you by
# Name = win32api.GetComputerName()


# create a function for the text to speech engine
def speak(audio1):
    engine.say(audio1)
    engine.runAndWait()


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


def speech_recognizer():
    # initialize KaldiRecognizer with the vosk model
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
            text_msg = text_msg1[14:-3]
            if len(text_msg) > 3:
                if 'what' or 'who' or 'when' or 'where' in text_msg:
                    new = str(text_msg) + '?'
                    print(new)
                    resp = generate_response(new)
                    speak(resp)
                else:
                    print(text_msg)
                    resp = generate_response(text_msg)
                    speak(resp)


if __name__ == '__main__':
    speech_recognizer()
