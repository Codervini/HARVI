import pyttsx3
import speech_recognition as sr
import datetime
import webbrowser
import os
import smtplib
import harvy_tasks as ht


MASTER = "Vinish"
VOICE = 0

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

    PROMPT = "I am Harvy a replica of jarvis. How may I help you?"

    if hour >=0 and hour < 12:
        speak("Good Morning" + MASTER + "  " + PROMPT)
    elif hour >= 12 and hour < 18:
        speak("Good Afternoon" + MASTER + "  " + PROMPT)
    else:
        speak("Good Evening" + MASTER + "  " + PROMPT)

    #speak("I am Harvy a replica of jarvis. How may I help you?") 

# Takes command from the microphone(STT)
r = sr.Recognizer()
class Command():
    def __init__(self):
        self
    def mic(self):
        with sr.Microphone() as source:
            speak("Initializing Jarvis...")
            r.adjust_for_ambient_noise(source, 5) 
            speak("Initialized Jarvis...")
            print("Listening.....")
            audio = r.listen(source)
        try:
            print("Recognizing......")
            query = r.recognize_google(audio, language= "en-in")
            print(f"User said: {query}\n")
        except Exception as e:
            print("Say that again please.....")
            query = 'None'
        return query.lower()
        
    def mic_supplement(self):
        with sr.Microphone() as source:
            print("Listening.....")
            audio = r.listen(source)
        try:
            print("Recognizing......")
            query = r.recognize_google(audio, language= "en-in")
            print(f"User said: {query}\n")
        except Exception as e:
            print("Say that again please.....")
            query = 'None'
        return query.lower()

# # # Function that uses SMTP to send email
# # def sendEmail(to, content):
# #     server = smtplib.SMTP('smtp.gmail.com', 587)
# #     server.ehlo()
# #     server.starttls()
# #     server.login('youremail@gmail.com', 'your-password')
# #     server.sendmail('youremail@gmail.com', to, content)
# #     server.close()

if __name__ == "__main__":
    wishMe()
    calibrate = True  

    while True: #Logic for executing tasks
        if calibrate == True:
            print("TT")
            query = Command().mic()
            calibrate = False
        else:
            query = Command().mic_supplement()

        if query != 'None':
            if 'wikipedia' in query:
                speak("Searching Wikipedia.....")
                speak(ht.Tasks(query).wikipedia1())
            elif 'open youtube' in query:
                url = "youtube.com"
                browser_path = 'C:/Program Files/Mozilla Firefox/firefox.exe %s'
                webbrowser.get(browser_path).open(url)
            elif 'open google' in query:
                url = "google.com"
                browser_path = 'C:/Program Files/Mozilla Firefox/firefox.exe %s'
                webbrowser.get(browser_path).open(url)

            elif 'play music' in query:
                songs_dir = "C:\\Users\\Vinish\\Desktop\\MUSIC BGM"
                songs = os.listdir(songs_dir)
                print(songs)
                os.startfile(os.path.join(songs_dir, songs[0]))

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



            # # elif 'email to harry' in query:
            # #             try:
            # #                 speak("What should I say?")
            # #                 content = takeCommand()
            # #                 to = "harryyourEmail@gmail.com"    
            # #                 sendEmail(to, content)
            # #                 speak("Email has been sent!")
            # #             except Exception as e:
            # #                 print(e)
            # #                 speak("Sorry my friend harry bhai. I am not able to send this email") 
        else:
            continue
