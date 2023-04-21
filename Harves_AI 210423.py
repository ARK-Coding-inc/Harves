#Imports
import pyttsx3
import speech_recognition as sr
import datetime
import time
from openpyxl import *
import random

#Variables
r = sr.Recognizer()
wakeWords = [("Jarvis", 1), ("Hey Jarvis", 1)]
source = sr.Microphone()

#Functions
def Speak(text):
    rate = 100
    engine = pyttsx3.init()
    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[0].id)
    engine.setProperty('rate', rate+50)
    engine.say(text)
    engine.runAndWait()

def callback(recognizer, audio):
    try:
        speech_as_text = recognizer.recognize_sphinx(audio, keyword_entries = wakeWords)
        print(speech_as_text)
        if "Jarvis" in speech_as_text or "Hey Jarvis" in speech_as_text:
            Speak("Yes?")
            recognize_main()
    except sr.UnknownValueError:
        print("Sorry! Didn't get that")
        
def start_recognizer():
    print("Waiting for a wake word... try 'Jarvis' or 'Hey Jarvis'")
    r.listen_in_background(source, callback)
    time.sleep(1000000)

def recognize_main():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        Speak("What can I help you with?")
        audio = r.listen(source)
        data = ""
        try:
            data = r.recognize_google(audio)
            data.lower()
            Speak(f"You said: {data}")
#Greetings-----------------------------------------------------------------------------------------------------
            #if hello in data
            if data in hello_list:
                hour = datetime.datetime.now().hour
                if hour >= 0 and hour < 12:
                    Speak("Good morning!")
                elif hour >= 12 and hour < 18:
                    Speak("Good afternoon!")
                else:
                    Speak("Good evening!")
            elif data in how_are_you:
                Speak(random.choice(reply_how_are_you))
            elif "What is the time" in data:
                strTime = datetime.datetime.now().strftime("%H:%M")
                Speak(f"The time is {strTime}")
            elif "what day is it" in data:
                day = datetime.datetime.today().weekday() + 1
                Day_dict = {1: 'Monday', 2: 'Tuesday', 3: 'Wednesday', 4: 'Thursday', 5: 'Friday', 6: 'Saturday', 7: 'Sunday'}
                if day in Day_dict.keys():
                    day_of_the_week = Day_dict[day]
                    print(day_of_the_week)
                    Speak(f"The day is day_of_the_week")
                    time.sleep(2)
            else:
                Speak("Sorry, I didn't understand your request")
        except sr.UnknownValueError:
            print("Jarvis did not understand your request")
        except sr.RequestError as e:
            print("Could not request results from Google Speech Recognition service: {0}".format(e))

def excel():
    wb = load_workbook("input.xlsx") #Opens Excel doc for data
    wu =  wb.get_sheet_by_name('User') #Sets sheet in Excel for user prompts
    wr = wb.get_sheet_by_name('Replies') #Setts sheet in Excel for replies

    global hello_list
    global how_are_you
    urow1 = wu['1']
    urow2 = wu['2']
    hello_list = [urow1[x].value for x in range(len(urow1))]
    how_are_you = [urow2[x].value for x in range(len(urow2))]

    global reply_hello_list
    global reply_how_are_you
    rrow1 = wr['1']
    rrow2 = wr['2']
    reply_hello_list = [rrow1[x].value for x in range(len(rrow1))]
    reply_how_are_you = [rrow2[x].value for x in range(len(rrow2))]

#Main Program
excel()
while 1:
    start_recognizer()