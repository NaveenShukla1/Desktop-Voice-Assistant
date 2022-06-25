import pyttsx3
import time
from win32com.client import Dispatch
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import random
import pyautogui
import sys
import instaloader


engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices)
engine.setProperty('voice', voices[0].id)

def speak(audio):

    engine.say(audio)
    engine.runAndWait()

#     speak_test = Dispatch("SAPI.spVoice")           
#     speak_test.Speak(audio)

def takeCommand():

    print("For exiting or leaving please speak:  'exit'")
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

        try: 
            print("Recogning...")
            query = r.recognize_google(audio, language='en-in')
            print(f"User said: {query}")

        except Exception:
            print("say that again please")
            return "None"
        return query

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour<12:
        speak("Good Morning")

    elif hour>=12 and hour<17:
        speak("Good Afternoon")

    else:
        speak("Good Evening")

    speak("I am jarvis. How may i help you! ")

def AllCommands():
    
    while True:
        query = takeCommand().lower()

        #logic for Executing tasks based on query
        try:

            if 'wikipedia' in query:
                speak('searching wikipedia...')
                query = query.replace("wikipedia", "")
                results = wikipedia.summary(query, sentences=1)
                speak("According to wikipedia")
                print(results)
                webbrowser.open(query)
                speak(results)
                return
        
            elif 'open youtube' in query:
                speak("Opening")
                webbrowser.open("Youtube.com")
                return

            elif 'open google' in query:
                speak("Opening")
                webbrowser.open("Google.com")
                return

            elif 'open stackoverflow' in query:
                speak("Opening")
                webbrowser.open("stackoverflow.com")
                return
            
            elif 'play music' in query:
                music_dir = 'C:\\Users\\Dell\\Music\\musics'
                songs = os.listdir(music_dir)
                #  os.startfile(os.path.join(directory, song name))
                os.startfile(os.path.join(music_dir, songs[random.randint(1,100)]))
                return

            elif 'time' in query:
                timeStr = datetime.datetime.now().strftime("%H:%M:%S")
                speak(f"sir, The time is {timeStr}")

            elif 'date' in query:
                dateStr = (datetime.datetime.now().day)
                speak(dateStr)

            elif 'month' in query:
                monthStr = (datetime.datetime.now().month)
                speak(monthStr)

            elif 'open code' in query:
                codePath = "C:\\Users\\Dell\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
                os.startfile(codePath)

            # elif 'open notepad' in query:
            #     codePath = "C:\\Users\\Dell\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            #     os.startfile(codePath)

            elif 'ThankYou' in query:
                speak("Your Welcome sir")

            elif 'exit' in query:
                speak("Goodbye")
                # exit()
                break

            elif 'switch windows' in query:
                speak('done')
                pyautogui.keyDown("alt")
                pyautogui.press("tab")
                pyautogui.keyUp("alt")


            elif "Instagram" in query or "instagram profile" in query or "profile on instagram" in query:
                speak("sir please enter the user name:")
                name = input("Enter username here: ")
                webbrowser.open(f"www.instagram.com/{name}")
                speak(f"sir, here is the profile of username {name}")
                time.sleep(2)
                speak("sir, would you like to download the profile picture of this account.")
                condition = takeCommand().lower()
                print(f"user said: {condition}")
                if 'yes' in condition:
                    mod = instaloader.Instaloader()
                    mod.download_profile(name, profile_pic_only= True)
                    speak("done sir, profile picture is saved in our main folder.")
                else:
                    pass
                return

            elif 'volume up' in query:
                pyautogui.press("volumeup")

            elif 'volume down' in query:
                pyautogui.press("volumedown")

            elif 'volume mute' in query or 'mute the volume' in query:
                pyautogui.press("volumemute")

            elif "screeenshot" in query:
                speak("sir, tell me the name for this screenshot file")
                name = takeCommand().lower()
                speak("Sir, please hold the screen for few seconds, i am taking a screenshot")
                img = pyautogui.screenshot()
                img.save(f"{name}.png")
                speak("done sir, the screenshot is saved in the main folder.")

            elif None in query:
                speak("Any task for me sir")
                condition2 = takeCommand().lower()
                if 'No' in condition2:
                    sys.exit()
                else:
                    AllCommands()



            # elif '' in query:
            #     if '' != 'exit' and  '' != 'Thankyou' :
            #         # query = query.replace("wikipedia", "")
            #         results = wikipedia.summary(query, sentences=2)
            #         speak("According to wikipedia")
            #         print(results)
            #         speak(results)
                    # print("if you want to search anything please say: ______according to wikipedia")

        except Exception:
            print("Try again")

if __name__ == '__main__':
    wishMe()

    AllCommands()
    # while True:
    #     permission = takeCommand().lower()
    #     if "ok jarvis "in permission  or "Hello jarvis" in permission:
    #         print(f"user said: {permission}")
    #         AllCommands()

    #     elif "exit" or "goodbye" in permission:
    #         sys.exit()
        
    #     else:
    #         print("please say 'okay jarvis' or 'hello jarvis' to activate your assistant")
    #         print("Or say exit or goodbye to exit")
