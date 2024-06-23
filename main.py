import pyttsx3 
import speech_recognition as sr
import os
import datetime
import webbrowser
from comtypes import CLSCTX_ALL
from ctypes import cast, POINTER
import instaloader
import requests
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import pyautogui as pau
import psutil
from time import sleep

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

def speak(txt):
    engine.say(txt)
    engine.runAndWait()

def news():
    main_url = "https://newsapi.org/v2/top-headlines?sources=techcrunch&apiKey=09d10de5c8e54f8b82c9e882d7aa4ff1"
    main_page = requests.get(main_url).json()
    articles = main_page["articles"]
    head =[]
    day = ["first", "second" ,"third", "fourth", "fifth", "sixth", "seventh", "eighth", "nineth", "tenth"]
    for ar in articles:
        head.append(ar["title"])

    for i in range(len(day)):
        print(f"Today's {day[i]} News is : \n {head[i]}")
        speak(f"Today's {day[i]} News is :{head[i]}")

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...") 
        r.pause_threshold = 1
        audio = r.listen(source)
        

    try:
        print("Recognising...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said : {query}")

    except Exception as e:
        print("Say that again please...")
        speak("Say that again please")

        return "none"
    return query

def count():
    speak('5')
    speak('4')
    speak('3')
    speak('2')
    speak('1')

def count2():
    speak('3')
    speak('2')

def half_volume():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    current_volume = volume.GetMasterVolumeLevelScalar()
    new_volume = current_volume / 2.0
    volume.SetMasterVolumeLevelScalar(new_volume, None)
    print("The Current Volume is :", new_volume)
    speak(f'The Current Volume is {new_volume}')

def full_volume():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume_range = volume.GetVolumeRange()
    max_volume_db = volume_range[1]
    volume.SetMasterVolumeLevel(max_volume_db, None)
    print("The Current Volume is 100%")
    speak("The Current Volume is 100%")

def wishMe():
    hour = int(datetime.datetime.now().hour)

    if hour>=0 and hour<=12:
       
        print("Good Morning Sir...")
        speak("Good Morning Sir!")
    
    elif hour>=12 and hour<=16:
       
        print("Good AfterNoon Sir...")
        speak("Good Afternoon Sir!")

    else:
        print("Good Evening Sir!")
        speak("Good Evening sir!")

    speak("I am jarvis., just a rather very important person. Please tell me how can i help you")
    sleep(1)

def main():
    while True:
        query = takeCommand().lower()
        if "hello" in query:
            speak("Hello sir, how can i help you today?")

        elif "insta mode" in query:
            speak('Wait for 3 Seconds')
            count2()
            speak('1')
            webbrowser.open('https://instagram.com/reels/')
            speak("sir speak down or up to down the reel section or up the reel section in instgram reels")
            while True:
                query = takeCommand().lower()
                if "down" in query:
                    pau.press('down')
                elif "up" in query:
                    pau.press('up')
                
                else:
                    pass
            
        elif"wish" in query:
            wishMe()
        
        elif"trades" in query:
            speak("okay sir which stock do you want to open?")
            make = takeCommand().lower()   
            make.replace(' ','')
            make.upper()
            print(make)
            webbrowser.open(f'https://in.tradingview.com/chart/?symbol=NSE%3A{make}')
            speak(f"the {make} stock opens in trading view site")
            print('Do you want to take screenshot?')
            speak('do you want to take screenshot')
            make_decision = takeCommand().lower()        
            if "yes" in make_decision :
                speak("Sir I am taking the screenshot dont move the screen for few seconds")
                print('But sir please tell me the name of the image file')
                file_name = takeCommand().lower()
                sleep(2)
                img = pau.screenshot()
                img.save(f'{file_name}.png')
                speak('done sir')
                print(f"Your screenshot has been saved by the name of {file_name}")
                speak(f"Your screenshot has been saved by the name of {file_name}")
            else:

                pass

        elif"switch the window" in query:
            print("Switching the Window in 3 seconds...")
            speak("Switching the window in 3 seconds")
            sleep(1)
            pau.keyDown("alt")
            pau.press("tab")
            sleep(1)
            pau.keyUp("alt")
        
        elif"insta profile"in query:
            print("Enter the username correctly...")
            speak("Enter the username correctly")
            name = input("Enter the Username : @")
            webbrowser.open(f"https://instagram.com/{name}")
            speak(f"Sir here is the profile of {name}")
            sleep(2)
            speak(f"Sir you want to download the profile picture of {name}")
            con = takeCommand().lower().replace("yash","yes").replace("yas","yes")
            if "yes" in con:
                mod = instaloader.Instaloader()
                mod.download_profile(name, profile_pic_only=True)
                speak("Your pic is ready and downloaded in main folder")
                pass

            else:
                pass
        
        elif "news" in query:
            print("Please Wait for Sometime.I am Fetching the Latest news...")
            speak("Please wait for some time. I am fetching the latest news.")
            news()
        elif"volume up" in query:
            pau.press("volumeup")

        elif"half the volume" in query:
            half_volume()
        
        elif"full the volume" in query:
            full_volume()
        
        elif"turn on the volume" in query:
            full_volume()
        
        elif"volume down" in query:
            pau.press("volumedown")
        
        elif"mute" in query:
            pau.press("volumemute")
        
        elif"unmute" in query:
            pau.press("volumemute")
        
        elif "battery" in query:
            print("Wait Sir1 Let me check...")
            speak("wait sir, let me check")
            battery = psutil.sensors_battery()
            percentage = battery.percent
            print(f"The System has {percentage}% battery")
            speak(f"The system has {percentage} percent battery")
        
        elif "shutdown" in query:
            print("The PC will be Shut Down in 5 seconds...")
            speak("The PC will be shut down in 5 seconds")
            count2()
            os.system("shutdown /s /t 1")

        elif "shut down" in query:
            print("The PC will be Shut Down in 5 seconds...")
            speak("The PC will be shut down in 5 seconds")
            count2()
            os.system("shutdown /s /t 1")

        elif "restart" in query:
            print("The PC will be restart in 5 seconds...")
            speak("The PC will be  restart in 5 seconds")
            count2()
            os.system("shutdown /r /t 1")

        elif"sleep the system" in query:
            print("The PC will be Gone in Sleep Mode in 5 Seconds...")
            speak("The pc will be gone in sleep mode in 5 seconds")
            count2()
            os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
        
        elif "leave" in query:
            print("Closing all the systems")
            speak("Closing all the systems")
            print("Nice to meet you sir")
            speak("Nice to meet you sir")
            os.system('taskkill /f /im code.exe')

     
if __name__ == '__main__':
    wishMe()
    main()