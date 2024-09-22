import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import os
import random

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)
engine.setProperty('rate', 150)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good morning Sir!")
    elif hour >= 12 and hour < 18:
        speak("Good afternoon Sir!")
    else:
        speak("Good evening Sir!")
    speak("Hi, I am Veronika. How can I help you Sir?")


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")
    except Exception as e:
        print("Could not understand the audio. Error:", e)
        print("Say that again please...")
        return "None"

    return query


def openApplication(app_name):
    if 'notepad' in app_name:
        speak("Opening Notepad")
        os.startfile("notepad.exe")

    elif 'excel' in app_name:
        speak("Opening Excel")
        codePath = "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"
        os.startfile(codePath)

    elif 'word' in app_name:
        speak("Opening Word")
        codePath = "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE"
        os.startfile(codePath)

    elif 'powerpoint' in app_name:
        speak("Opening PowerPoint")
        codePath = "C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE"
        os.startfile(codePath)

    elif 'telegram' in app_name:
        speak("Opening Telegram")
        codePath = "C:\\Users\\x\\AppData\\Roaming\\Telegram Desktop\\Telegram.exe"
        os.startfile(codePath)

    elif 'chrome' in app_name:
        speak("Opening Chrome")
        codePath = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
        os.startfile(codePath)

    elif 'paint' in app_name:
        speak("Opening Paint")
        os.startfile("mspaint.exe")

    else:
        speak("Application not recognized or not available")


if __name__ == '__main__':
    wishMe()

    while True:
        query = takeCommand().lower()

        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            speak("Opening YouTube")
            webbrowser.open("https://www.youtube.com")

        elif 'open google' in query:
            speak("Opening Google")
            webbrowser.open("https://www.google.co.in/")

        elif 'open chat gpt' in query:
            speak("Opening Chat GPT")
            webbrowser.open("https://chat.openai.com/")

        elif 'play music' in query or 'open music'in query or 'play songs'in query:
            music_dir = 'C:\\Users\\x\\Desktop\\veronika\\my music'
            songs = os.listdir(music_dir)
            if songs:
                random_song = random.choice(songs)
                print(f"Playing: {random_song}")
                os.startfile(os.path.join(music_dir, random_song))
            else:
                speak("No songs found in the directory")


        elif 'open wallpaper' in query:
            wallpaper_dir = 'C:\\Users\\x\\Desktop\\veronika\\jpg wallpapers'
            wall = os.listdir(wallpaper_dir)
            if wall:
                random_wall = random.choice(wall)
                print(f"Playing: {random_wall}")
                os.startfile(os.path.join(wallpaper_dir, random_wall))
            else:
                speak("No songs found in the directory")
        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {strTime}")

        elif 'open' in query:
            app_name = query.replace('open', '').strip()
            openApplication(app_name)

        elif 'exit' in query or 'stop' in query:
            speak("Goodbye Sir, have a nice day!")
            break

        else:
            speak("Sorry, I didn't catch that. Can you please repeat?")
