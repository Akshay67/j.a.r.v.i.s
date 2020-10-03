from win32com.client import Dispatch # pip install -m pip install pywin32
import pyttsx3 # pip install pyttsx3
import speech_recognition as sr # pip install SpeechRecognition
import datetime
import wikipedia # pip install wikipedia
import webbrowser # pip install pip install webbrowser
import os
import random
import smtplib # pip install smtplib
import pytz # pip install pytz
import pyautogui # pip install pyautogui
import psutil # pip install psutil
import pyjokes # pip install pyjokes
import sys
import wolframalpha

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def hello():
    speak('hello sir.. how are you ?')
    reply = takeCommand()
    if 'fine' in reply :
        speak('ok! thats great!.. ')
    elif 'not good' in reply or 'not well' in reply or 'ill' in reply:
        speak('please take care sir.. and please have some rest..')

    speak('do u need anything from me.. ?')
    reply3 = takeCommand()
    if 'yes' in reply3 :
        speak('how can i help you ?')

    elif 'no' in reply3 :
        speak('ok sir.. no problem.. remind me.. if you need anything from me.. thank you..')
        sys.exit()

    else:
        speak ('sir  please tell me.. how may i help you?  ')



def time():
    t_now = datetime.datetime.now().strftime('%H:%M:%S')
    speak("sir. the current time is")
    print(t_now)
    speak(t_now)
    speak('what should i do next sir?')



def greet():
    t_hour = datetime.datetime.now().hour

    if 24 > t_hour < 4:
        speak('pleasant night sir..')

    elif 4 <= t_hour < 12 :
        speak("good morning sir.. have a nice day..")

    elif 12 <= t_hour < 17 :
        speak("good afternoon sir.. hope u enjoying ur day")

    elif 17 <= t_hour < 19:
        speak('good evening sir.. hope u enjoying ur day')

    else :
        speak('good night sir.. hope u enjoyed ur day')      


    speak("jarvis ata your service sir..")
    speak('command me! sir!')

def date():
    t_date = datetime.datetime.now( tz = pytz.timezone('Asia/Kolkata'))
    speak('todays date is')
    print(t_date.strftime('%d %B, of %Y'))
    speak(t_date.strftime('%d %B, of %Y'))
    speak('Next Command! Sir!')


def note():
    speak ('what should i note down ?')
    content = takeCommand()
    speak('okay')
    remember = open("notes.txt",'w')
    remember.write(content)
    remember.close()
    speak('i have taken a note')
    speak('next command! sir!')



def chrome():
    speak("what should i search ?")
    search = takeCommand().lower()
    speak('okay! searching! ')
    chromepath = 'C://Program Files (x86)//Google//Chrome//Application//chrome.exe %s'
    webbrowser.get(chromepath).open_new_tab('www.google.com/#q='+search)
    speak('Next Command! Sir!')

def screenshot():
    speak('okay')
    img = pyautogui.screenshot()
    img.save('C:\\Users\\user\\Desktop\\jarvis\\screenshots\\ss.png')
    speak('i have taken screenshot')
    speak('Next Command! Sir!')

    

def jokes():
    my_joke = pyjokes.get_joke('en',category='neutral')
    print(my_joke)
    speak(my_joke)
    speak('Next Command! Sir!')



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

        speak('pardon sir..')
        print("Say that again please...")
        return "None"
    return query


if __name__ == "__main__":
    greet()
    while True:

        query = takeCommand().lower()

        if 'open youtube' in query:
            speak('okay opening youtube')
            webbrowser.open_new_tab("youtube.com")
            speak('Next Command! Sir!')
            

        elif 'open google' in query:
            speak('okay opening google')
            webbrowser.open_new_tab("google.com")
            speak('Next Command! Sir!')

        elif 'open gmail' in query:
            speak('okay opening gmail')
            webbrowser.open_new_tab('www.gmail.com')
            speak('Next Command! Sir!')

        elif 'time' in query:
            time() 

        elif 'search on chrome' in query or 'search on google' in query or 'chrome search' in query or 'google search' in query:
            chrome()


        elif 'date' in query :
            date()


        elif 'shutdown' in query :
            speak ('do u really want to shutdown ?')
            reply = takeCommand()

            if 'yes' in reply :
                speak(' bye sir.. i am shutting down our system.. take care sir.. thank you..')
                os.system('shutdown /s /t 1')

            elif 'no' in reply:
                speak('ok no problem sir. please proceed ur work')
                speak('tell me sir! what can i do for u next?  ')

            else :
                speak('please tell me.. how may i help you?  ')

        elif 'restart' in query :
            speak ('do u really want to restart ?')
            reply = takeCommand()

            if 'yes' in reply :
                speak('okay')
                os.system('shutdown /r /t 1')

            elif 'no' in reply :
                speak('ok no problem sir. please proceed ur work')
                speak('tell me sir! what can i do for u next?  ')

            else :
                speak('please tell me  how may i help you?  ')

        elif 'joke' in query:
            speak('okay i am telling a joke! ')
            jokes()


        elif 'according to google' in query:
            speak('okay sir! searching! ')
            chromepath = 'C://Program Files (x86)//Google//Chrome//Application//chrome.exe %s'
            query = query.replace('according to google','')
            results = wikipedia.summary(query, sentences = 3)
            speak('Got it sir! ')     
            speak('google says - ')
            print(results)
            webbrowser.get(chromepath).open_new_tab('www.google.com/#q='+query)
            speak(results)
            speak(f"sir! i also opened result on google! if you want extra information about {query}.. you can visit there!")
            speak('Next Command! Sir!')
            

        elif 'exit' in query or 'quit' in query or 'abort' in query or 'stop' in query or 'bye' in query:
            speak('bye sir. remind me.. if u need anything from me. thank you..')
            break


        elif 'hello' in query or 'hi' in query :
            hello()


        elif 'note' in query or 'notes' in query :
            note()

        elif 'how are you' in query:
            stMsgs = ['Just doing my thing sir..!  ', 'I am fine and full of energy!', 'I am nice and full of energy']
            speak(random.choice(stMsgs))
            speak(' how are you? ')
            reply = takeCommand()
            if 'fine' in reply :
                speak('ok! thats great! ')
            elif 'not good' in reply or 'not well' in reply or 'ill' in reply:
                speak('please take care sir.. and please have some rest..')


            speak('do u need anything from me.. ?')
            reply3 = takeCommand()
            if 'yes' in reply3 :
                speak('how can i help you ?')

            elif 'no' in reply3 :
                speak('ok sir.. no problem.. remind me.. if you need anything from me.. thank you  ')
                sys.exit()

            else:
                speak ('sir  please tell me.. how may i help you?  ')





        elif 'take screenshot' in query :
            screenshot()
            
        elif 'thank you' in query or 'thanks' in query :
            speak('sir! You are embarrassing me!. I am just doing my job!')
            speak('Next Command! Sir!')
            

        else:
            query = query
            speak('Searching...')
            try:
                try:
                    res = client.query(query)
                    results = next(res.results).text
                    speak('WOLFRAM-ALPHA says - ')
                    speak('Got it sir!')
                    speak(results)
                    print(results)

                except:
                    results = wikipedia.summary(query, sentences = 2)
                    speak('Got it sir! ')
                    speak('WIKIPEDIA says - ')
                    print(results)
                    speak(results)


            except:
                speak('sorry sir! i didnt get results..! please search on google!')
                speak('i am opening google! thank you !')
                webbrowser.open('www.google.com')

            speak('Next Command! Sir!')
