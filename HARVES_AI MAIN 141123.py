#imports
import pyttsx3
import time
import speech_recognition as sr
import random
import datetime
from openpyxl import *

#variable defs
r = sr.Recognizer()
keywords = [("Harves",1),("Hey Harves",1)]
source = sr.Microphone()

#function defs
def Speak(text):
    rate = 100
    engine = pyttsx3.init()
    voices = engine.getProperty('voices')
    engine.setProperty('voice',voices[0].id)
    engine.setProperty('rate', rate+50)
    engine.say(text)
    engine.runandWait()
def callback(recognizer, audio):
    try: 
        speech_as_text = recognizer.recognize_sphinx(audio, keyword_entries=keywords)
        print(speech_as_text)
        if "Harves" in speech_as_text or "Hey Harves":
            Speak("Yes sir?")
            recognize_main()
    except sr.UnknownValueError:
        print("Oops! Didn't catch that!")
def start_recognizer():
    print("Waiting for a keyword... 'Harves' or 'Hey Harves'")
    r.listen_in_background(source, callback)
    time.sleep(1000000)
def recognize_main():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("What can I do for you?")
        audio = r.listen(source)
    data = ""
    try:
        data = r.recognize_google(audio)
        data.lower()
        print("You said: "+ data)
#Greetings--------------------------------------------------------------------------------------
        if data in hello_list:
            hour = datetime.datetime.now().hour
            if hour >= 0 and hour < 12:
                Speak("Good morning, captain")
            elif hour >= 12 and hour < 18:
                Speak("Good afternoon, captain")
            else:
                Speak("Good evening, captain")
            Speak(random.choice(reply_hello_list))
            time.sleep(2)
        elif data in how_are_you:
            if hour >= 0 and hour < 12:
                Speak("Good morning, captain")
            elif hour >= 12 and hour < 18:
                Speak("Good afternoon, captain")
            else:
                Speak("Good evening, captain")
            Speak(random.choice(reply_how_are_you))
            time.sleep(2)
        elif "What is the time" in data:
            strTime = datetime.datetime.now().strftime("%H:%M")
            Speak(f"The time is {strTime}")
            time.sleep(2)
        elif "What day is it" in data:
            day = datetime.datetime.today().weekday() + 1
            Day_dict = {1: 'Monday', 2: 'Tuesday', 3: 'Wednesday',
                        4: 'Thursday', 5: 'Friday', 6: 'Saturday',
                        7: 'Sunday'}
            if day in Day_dict.keys():
                day_of_the_week = Day_dict[day]
                Speak("The day is " + day_of_the_week)
            time.sleep(2)  
        else:
            Speak("Sorry captain, I couldn't understand that")
    except sr.UnknownValueError:
        print("Harves did not understand your request")
    except sr.RequestError as e:
        print("Could not request results from Google Speech Recognition Service; {0}".format(e))
def excel():
    wb = load_workbook("input.xlsx")
    wu = wb.get_sheet_by_name('User')
    wr = wb.get_sheet_by_name('Replies')

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

#main command
excel()
while 1:
    start_recognizer()