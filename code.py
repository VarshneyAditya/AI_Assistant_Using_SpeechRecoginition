import pyttsx3 as textToSpeach                 # pip install pyttsx3
#import pip install pypiwin32.                 # use this library if pyttsx3 get error like No module named win32com. client No module named win32. No module named win32ap
import SpeechRecognition as speechRecognition # pip install speechRecognition
import datetime as date                        # work with dates
import wikipedia as wiki                       # pip install wikipedia . For Wikipedia searches
import webbrowser as browser                   # For open any website
import os as playMusic                         # for play music
import smtplib as sendmail                     # for send mail and rcieve mail

engine = textToSpeach.init('sapi5')            # sapi5 => Helps in synthesis and recognition of voice
voices = engine.getProperty('voices')          # voice => getting details of current voice
# print(voices[1].id)
engine.setProperty('voice', voices[0].id)


def speak(audio):                              # speak() function to convert our text to speech.
    engine.say(audio)
    engine.runAndWait()                        # Without this command, speech will not be audible to us.


def wishMe():                                  # wishme() function that gives the greeting functionality according to our A.I system time
    hour = int(date.date.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!")
    elif hour>=12 and hour<18:
        speak("Good Afternoon!")   
    else:
        speak("Good Evening!")  
    speak("I am Jarvis Sir. Please tell me how may I help you")       


def takeCommand(): #takeCommand() function, which helps our A.I to take command from the user. It takes microphone input from the user and returns string output.

    r = speechRecognition.Recognizer()           
    with speechRecognition.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")    
        query = r.recognize_google(audio, language='en-in')     #Using google for voice recognition.
        print(f"User said: {query}\n")                          #User query will be printed.

    except Exception as e:
        # print(e)    
        print("Say that again please...")                       #Say that again will be printed in case of improper voice
        return "None"                                           #None string will be returned
    return query

def sendEmail(to, content):                                     # send emails to one or more than one recipient
    server = sendmail.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('youremail@gmail.com', 'your-password')
    server.sendmail('youremail@gmail.com', to, content)
    server.close()

if __name__ == "__main__":
    wishMe()
    while True:
    # if 1:
        query = takeCommand().lower()                           #Converting user query into lower case

 # Logic for executing tasks based on query
        if 'wiki' in query:                                     #if wikipedia found in the query then this block will be executed
            speak('Searching wiki...')
            query = query.replace("wiki", "")
            results = wiki.summary(query, sentences=2)
            speak("According to wiki")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            browser.open("youtube.com")                         # open youtube on browser by using browser.open

        elif 'open google' in query:
            browser.open("google.com")

        elif 'open stackoverflow' in query:
            browser.open("stackoverflow.com")   


        elif 'play music' in query:
            music_dir = 'Text' #Write your own directory path where music files are stored.
            songs = playMusic.listdir(music_dir)
            print(songs)    
            playMusic.startfile(playMusic.path.join(music_dir, songs[0]))

        elif 'the time' in query:
            strTime = date.date.now().strftime("%H:%M:%S")    
            speak(f"Sir, the time is {strTime}")

        elif 'open code' in query:
            codePath = "C:\\Users\\[Mention your user name here]\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            playMusic.startfile(codePath)

        elif 'email to reciepent' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                to = "reciepent email Id"    
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")   