import pyttsx3
import speech_recognition as sr
import datetime
import webbrowser
import os
import harvy_tasks as ht


MASTER = "Vinish"
VOICE = 0
SONGS_DIR = "C:\\Users\\Vinish\\Desktop\\MUSIC BGM"
BROWSER_PATH = 'C:/Program Files/Mozilla Firefox/firefox.exe %s'

#The instance which TTS function
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[VOICE].id)

# A function that Pronunces the string passed to it
def speak(text):
    engine.say(text)
    engine.runAndWait()

# A Function that wishes the user depending upon the time
def wishMe():
    hour = int(datetime.datetime.now().hour)

    PROMPT = "I am Harvi. How may I help you?"

    if hour >=0 and hour < 12:
        speak("Good Morning" + MASTER + "  " + PROMPT)
    elif hour >= 12 and hour < 18:
        speak("Good Afternoon" + MASTER + "  " + PROMPT)
    else:
        speak("Good Evening" + MASTER + "  " + PROMPT)

    #speak("I am Harvy a replica of jarvis. How may I help you?") 

# Takes command from the microphone(STT)
r = sr.Recognizer()
def recognizer(mic_input):
    print("Listening.....")
    audio = r.listen(mic_input)
    try:
        print("Recognizing......")
        query = r.recognize_google(audio, language= "en-in")
        print(f"User said: {query}\n")
    except Exception as e:
        print("Say that again please.....")
        query = 'None'
    return query.lower()

def is_calibrate(bool : bool):
    with sr.Microphone() as source:
        if bool == True:
            speak("Initializing Jarvis...")
            r.adjust_for_ambient_noise(source, 5) 
            speak("Initialized Jarvis...")
            return recognizer(source)
        elif bool == False:
            return recognizer(source)

if __name__ == "__main__":
    wishMe()
    calibrate = True  
    while True: #Logic for executing tasks
        if calibrate == True:
            query = is_calibrate(calibrate)
            calibrate = False
        else:
            query = is_calibrate(calibrate)
        print(query)

        if query != 'None':
            if 'wikipedia' in query:
                speak("Searching Wikipedia.....")
                speak(ht.wikipedia1(query))
            elif 'open youtube' in query:
                url = "youtube.com"
                browser_path = 'C:/Program Files/Mozilla Firefox/firefox.exe %s'
                webbrowser.get(browser_path).open(url)
            elif 'open google' in query:
                url = "google.com"
                browser_path = 'C:/Program Files/Mozilla Firefox/firefox.exe %s'
                webbrowser.get(browser_path).open(url)

            elif 'play music' in query:
                songs = os.listdir(SONGS_DIR)
                print(songs)
                os.startfile(os.path.join(SONGS_DIR, songs[0]))

            elif 'the time' in query:
                strTime = datetime.datetime.now().strftime("%H:%M:%S")
                speak(f'{MASTER} the time is {strTime}')

            elif 'open code' in query:
                codePath = 'C:\\Users\\Vinish\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe'
                os.startfile(codePath)
            elif 'bye' in query:
                print("Good Bye Sir")
                speak("Good Bye Sir")
                break
        else:
            continue
