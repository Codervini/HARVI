import pyttsx3
import speech_recognition as sr
import datetime
import webbrowser
import os
import task_manager as tmgr
from tdprocessor import taskwriter


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
def active_recognizer(mic_input):
    '''Returns text from the given audio argument. 
    Should be used only when HARVI should be executing tasks.'''
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

def passive_recognizer(mic_input):
    '''Used when HARVI shouldn't be executing any tasks'''
    print("Passive Mode")
    audio = r.listen(mic_input)
    try:
        print("Recognizing passively......")
        query = r.recognize_google(audio, language= "en-in")
    except Exception as e:
        query = 'None'
    return query.lower()



def mic(listen = True, calibrate = True):
    '''This function takes care of everything related to mic'''
    if listen == True:
        with sr.Microphone() as source:
            if calibrate == True:
                speak("Calibrating the mic")
                r.adjust_for_ambient_noise(source, 3) 
                speak("Ready to listen")
                return active_recognizer(source)
            elif calibrate == False:
                return active_recognizer(source)
    elif listen == False:
        with sr.Microphone() as source:
            return passive_recognizer(source)




if __name__ == "__main__":
    wishMe()
    listen = True
    calibrate = True  

    while True: #Logic for executing tasks
        if calibrate == True and listen == True:
            query = mic(calibrate=calibrate)
            calibrate = False
        else:
            query = mic(listen=listen, calibrate=calibrate)

        if query != 'None' and listen == True:
            if 'recalibrate mic' in query:
                mic(calibrate=True)
            elif 'wikipedia' in query:
                speak("Searching Wikipedia.....")
                speak(tmgr.wikipedia1(query))
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
            elif 'open app' in query:
                query = query [query.index("open app") + len("open app"): ].strip()
                print(query)
                if len(query) >= 1:
                    tmgr.openapp(query)
            elif 'add apps' in query:
                taskwriter()
            elif 'stop listening' in query:
                listen = False
            elif 'bye' in query:
                print("Good Bye Sir")
                speak("Good Bye Sir")
                break
        elif listen == False:
            if query in ['listen to me', 'start listening']:
                listen = True
        else:
            continue









   