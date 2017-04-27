import speech_recognition as sr
import win32com.client as wincl
import time
from time import ctime
import os
import webbrowser

def speak(soundTrack):
    speak = wincl.Dispatch("SAPI.SpVoice")
    speak.Speak(soundTrack)

## Open Apps by command 
def browser():
    os.system("start chrome.exe")
def music():
    os.system("start wmplayer.exe")

# close apps 
def musicClose():
    os.system("taskkill /IM wmplayer.exe /F")

def stopBrowser():
    os.system("taskkill /IM chrome.exe /F")

    
## take voice
def takeVoice():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source)
    data = " "
    try:
        data = r.recognize_google(audio)
        print('You said: ' + r.recognize_google(audio))
    except sr.UnknownValueError:
        print("I could not understand.")
        speak("Sorry,I could not understand. please Try again ,")
    except sr.RequestError as e:
        print("{0}".format(e))
    return data

def commandData(data):
    if v == 'who are you':
        speak("I am Vubon. I am a machine")
    elif v == "how are you":
        speak('I am fine.')
        
    elif v == 'Browser':
        speak('Opening browser')
        browser()
        
    elif v == 'close browser':
        speak("Close browser")
        stopBrowser()
        
    elif v == 'music':
        speak('Opening music player')
        music()
        
    elif v =='close music':
        speak('close music')
        musicClose()
        
    elif v == 'time':
        speak(ctime())
        
    elif v == 'Facebook':
        fb = str(v)
        speak('Hold on please , I am login facebook now .')
        url = 'https://facebook.com'
        webbrowser.open_new(url)

time.sleep(2)
speak("Sir, How can I help you?")

while 1:
    v = takeVoice()
    commandData(v)
    if v == "exit":
        break
