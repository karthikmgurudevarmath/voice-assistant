import subprocess
import wolframalpha
import pyttsx3
import tkinter
import json
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import feedparser
import smtplib
import ctypes
import time
import requests
import shutil
from twilio.rest import Client
from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices[1].id)
engine.setProperty('voice', voices[1].id)

def speak(audio):
    # Using PowerShell for robust Text-to-Speech
    try:
        # Escape single quotes for PowerShell
        safe_audio = audio.replace("'", "''")
        command = f"Add-Type -AssemblyName System.Speech; (New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak('{safe_audio}')"
        subprocess.call(["powershell", "-NoProfile", "-Command", command])
    except Exception as e:
        print("TTS Error:", e)
        # Fallback to print if even that fails
        print(f"Assistant: {audio}")

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning Sir !")

    elif hour >= 12 and hour < 18:
        speak("Good Afternoon Sir !")

    else:
        speak("Good Evening Sir !")

    speak("I am your Assistant")
    speak("How can I help you") 

def username():
    speak("What should i call you sir")
    uname = takeCommand()
    speak("Welcome Mister")
    speak(uname)
    columns = shutil.get_terminal_size().columns
    
    print("#####################".center(columns))
    print("Welcome Mr.", uname.center(columns))
    print("#####################".center(columns))
    
    speak("How can i Help you, Sir")

def takeCommand():
    # Adding a keyboard fallback for easier interaction
    print("\n[V] Listening for voice or [K] Type command...")
    
    r = sr.Recognizer()
    r.energy_threshold = 300  # Adjust for ambient noise
    r.dynamic_energy_threshold = True
    
    audio = None
    try:
        with sr.Microphone() as source:
            print("Listening...")
            r.adjust_for_ambient_noise(source, duration=0.5)
            
            # Popup logic
            popup = None
            try:
                popup = tkinter.Tk()
                popup.title("Jarvis")
                popup.overrideredirect(True)
                popup.attributes('-topmost', True)
                width, height = 300, 100
                screen_width = popup.winfo_screenwidth()
                screen_height = popup.winfo_screenheight()
                popup.geometry(f'{width}x{height}+{(screen_width//2)-(width//2)}+{screen_height-height-100}')
                popup.configure(bg='#202124')
                tkinter.Label(popup, text="🎤 Listening...", font=("Segoe UI", 16), fg="white", bg='#202124').pack(expand=True)
                popup.update()
            except Exception:
                pass

            try:
                # Increased timeout for real-world usage
                audio = r.listen(source, timeout=5, phrase_time_limit=10)
            except sr.WaitTimeoutError:
                print("Voice timeout: No speech detected.")
            except Exception as e:
                print(f"Voice capture error: {e}")
            
            if popup:
                try: popup.destroy()
                except: pass

        if audio:
            try:
                print("Recognizing...")
                query = r.recognize_google(audio, language='en-in')
                print(f"User said (Voice): {query}\n")
                return query
            except sr.UnknownValueError:
                print("Google Speech Recognition could not understand audio.")
            except sr.RequestError as e:
                print(f"Could not request results from Google Speech Recognition service; {e}")
            except Exception as e:
                print(f"Recognition Error: {e}")

    except Exception as e:
        # This usually means no microphone is found or permission denied
        print(f"Microphone Error: {e}")

    # Keyboard Fallback
    print("Fallback to Keyboard. Enter command: ", end="", flush=True)
    keyboard_query = input()
    if keyboard_query.strip():
        return keyboard_query
    
    return "None"

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    
    # Enable low security in gmail
    server.login('saanji31@gmail.com', 'xqos menp lduq caen')
    server.sendmail('saanji31@gmail.com', to, content)
    server.close()

if __name__ == '__main__':
    clear = lambda: os.system('cls')
    
    # This Function will clean any
    # command before execution of this python file
    clear()
    wishMe()
    username()
    
    while True:
        
        query = takeCommand().lower()
        
        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences = 3)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            speak("Here you go to Youtube\n")
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            speak("Here you go to Google\n")
            webbrowser.open("google.com")

        elif 'open stackoverflow' in query:
            speak("Here you go to Stack Over flow.Happy coding")
            webbrowser.open("stackoverflow.com")   

        elif 'play music' in query or "play song" in query:
            speak("Here you go with music")
            # music_dir = "G:\\Song"
            music_dir = "C:\\Users\\GAURAV\\Music" # Adjust this path
            try:
                songs = os.listdir(music_dir)
                print(songs)    
                random_song = os.path.join(music_dir, songs[1]) 
                os.startfile(random_song)
            except Exception as e:
                speak("Music directory not found")

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")    
            speak(f"Sir, the time is {strTime}")

        elif 'open opera' in query:
            # codePath = r"C:\\Users\\GAURAV\\AppData\\Local\\Programs\\Opera\\launcher.exe"
            # os.startfile(codePath)
            speak("Opening Opera")

        elif 'email to gaurav' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                to = "Receiver email address"    
                sendEmail(to, content)
                speak("Email has been sent !")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")

        elif 'send a mail' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                speak("whome should i send")
                to = input()    
                sendEmail(to, content)
                speak("Email has been sent !")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")

        elif 'how are you' in query:
            speak("I am fine, Thank you")
            speak("How are you, Sir")

        elif 'fine' in query or "good" in query:
            speak("It's good to know that your fine")

        elif "change my name to" in query:
            query = query.replace("change my name to", "")
            assname = query

        elif "change name" in query:
            speak("What would you like to call me, Sir ")
            assname = takeCommand()
            speak("Thanks for naming me")

        elif "what's your name" in query or "What is your name" in query:
            speak("My friends call me")
            speak("Jarvis") # Default name
            print("My friends call me Jarvis")

        elif 'exit' in query:
            speak("Thanks for giving me your time")
            exit()

        elif "who made you" in query or "who created you" in query: 
            speak("I have been created by Gaurav.")

        elif 'joke' in query:
            speak(pyjokes.get_joke())

        elif "calculate" in query: 
            
            app_id = "QHH44H6EE9" 
            indx = query.lower().split().index('calculate') 
            query_content = query.split()[indx + 1:] 
            string_query = ' '.join(query_content)
            
            # Using direct API to avoid library issues
            try:
                import urllib.parse
                encoded_query = urllib.parse.quote(string_query)
                url = f"http://api.wolframalpha.com/v1/result?appid={app_id}&i={encoded_query}"
                response = requests.get(url)
                
                if response.status_code == 200:
                    answer = response.text
                    print("The answer is " + answer) 
                    speak("The answer is " + answer)
                else:
                    speak("I couldn't calculate that.")
                    print("Error:", response.text)
            except Exception as e:
                print(e)
                speak("Something went wrong with the calculation") 

        elif 'search' in query or 'play' in query:
            
            query = query.replace("search", "") 
            query = query.replace("play", "")         
            webbrowser.open(query) 

        elif "who i am" in query:
            speak("If you talk then definitely your human.")

        elif "why you came to world" in query:
            speak("Thanks to Gaurav. further It's a secret")

        elif 'power point presentation' in query:
            speak("opening Power Point presentation")
            # power = r"C:\\Users\\GAURAV\\Desktop\\Minor Project\\Presentation\\Voice Assistant.pptx"
            # os.startfile(power)
            speak("Path to presentation usually goes here")

        elif 'is love' in query:
            speak("It is 7th sense that destroy all other senses")

        elif "who are you" in query:
            speak("I am your virtual assistant created by Gaurav")

        elif 'reason for you' in query:
            speak("I was created as a Minor project by Mister Gaurav ")

        elif 'change background' in query:
            # ctypes.windll.user32.SystemParametersInfoW(20, 0, "Location of wallpaper", 0)
            speak("Background change functionality is commented out for safety")

        elif 'open bluestack' in query:
            # appli = r"C:\\ProgramData\\BlueStacks\\Client\\Bluestacks.exe"
            # os.startfile(appli)
            speak("Opening Bluestacks")

        elif 'news' in query:
            try:
                jsonObj = urlopen(
                    'https://newsapi.org/v2/top-headlines?sources=the-times-of-india&apiKey=71801df9a62544189ce621f12c9d9d99'
                )
                data = json.load(jsonObj)
                i = 1

                speak('Here are some top news from the Times of India')
                print("=============== TIMES OF INDIA ===============\n")

                for item in data['articles']:
                    print(str(i) + '. ' + item['title'] + '\n')
                    speak(str(i) + '. ' + item['title'])
                    i += 1
            except Exception as e:
                print("News Error:", e)
                speak("Sorry, I am unable to fetch news right now")

        elif 'lock window' in query:
                speak("locking the device")
                ctypes.windll.user32.LockWorkStation()

        elif 'shutdown system' in query:
                speak("Hold On a Sec ! Your system is on its way to shut down")
                subprocess.call('shutdown /p /f')
                
        elif 'empty recycle bin' in query:
            try:
                winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
                speak("Recycle Bin Recycled")
            except Exception as e:
                print("Recycle bin error:", e)

        elif "don't listen" in query or "stop listening" in query:
            speak("for how much time you want to stop jarvis from listening commands")
            a = int(takeCommand())
            time.sleep(a)
            print(a)

        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.open("https://www.google.nl/maps/place/" + location + "")

        elif "camera" in query or "take a photo" in query:
            ec.capture(0, "Jarvis Camera ", "img.jpg")

        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])
            
        elif "hibernate" in query or "sleep" in query:
            speak("Hibernating")
            subprocess.call("shutdown /h")

        elif "log off" in query or "sign out" in query:
            speak("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])

        elif "write a note" in query:
            speak("What should i write, sir")
            note = takeCommand()
            file = open('jarvis.txt', 'w')
            speak("Sir, Should i include date and time")
            snfm = takeCommand()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("%H:%M:%S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)
        
        elif "show note" in query:
            speak("Showing Notes")
            file = open("jarvis.txt", "r") 
            print(file.read())
            speak(file.read(6))

        elif "update assistant" in query:
            speak("After downloading file please replace this file with the downloaded one")
            pass

        elif "jarvis" in query:
            wishMe()
            speak("Jarvis 1 point o in your service Mister")

        elif "weather" in query:
            # Google Open weather website
            # to get API of Open weather 
            api_key = "Api key" # Placeholder
            base_url = "http://api.openweathermap.org/data/2.5/weather?"
            speak(" City name ")
            print("City name : ")
            city_name = takeCommand()
            complete_url = base_url + "appid=" + api_key + "&q=" + city_name
            response = requests.get(complete_url) 
            x = response.json() 
            
            if x["code"] != "404": 
                y = x["main"] 
                current_temperature = y["temp"] 
                current_pressure = y["pressure"] 
                current_humidiy = y["humidity"] 
                z = x["weather"] 
                weather_description = z[0]["description"] 
                print(" Temperature (in kelvin unit) = " +
                                str(current_temperature) + 
                    "\n atmospheric pressure (in hPa unit) ="+
                                str(current_pressure) +
                    "\n humidity (in percentage) = " +
                                str(current_humidiy) +
                    "\n description = " +
                                str(weather_description)) 
            else: 
                speak(" City Not Found ")

        elif "send message " in query:
                # You need to create an account on Twilio to use this service
                account_sid = 'Account Sid key'
                auth_token = 'Auth token'
                client = Client(account_sid, auth_token)
    
                message = client.messages \
                                .create(
                                    body = takeCommand(),
                                    from_='Sender No',
                                    to ='Receiver No'
                                )
                print(message.sid)
    
        elif "wikipedia" in query:
            webbrowser.open("wikipedia.com")
    
        elif "Good Morning" in query:
            speak("A warm" +query)
            speak("How are you Mister")
    
        # most asked question from google Assistant
        elif "will you be my gf" in query or "will you be my bf" in query:   
            speak("I'm not sure about, may be you should give me some time")
    
        elif "how are you" in query:
            speak("I'm fine, glad you me that")
    
        elif "i love you" in query:
            speak("It's hard to understand")
    
        elif "what is" in query or "who is" in query:
            # Use the same API key 
            # Use the same API key 
            app_id = "QHH44H6EE9"
            
            try:
                import urllib.parse
                # Ensure we strip the trigger words if needed, but usually query is ok
                # logic to clean query if it contains "what is"
                clean_query = query.replace("what is", "").replace("who is", "")
                encoded_query = urllib.parse.quote(clean_query)
                url = f"http://api.wolframalpha.com/v1/result?appid={app_id}&i={encoded_query}"
                response = requests.get(url)
                
                if response.status_code == 200:
                    print(response.text)
                    speak(response.text)
                else:
                    print("No short answer available")
                    speak("I am not sure how to answer that")
            except Exception as e:
                 print (e)
                 speak("I encountered an error looking that up")
