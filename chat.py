import speech_recognition as sr
import pyttsx3 
import ctypes 
import time 
import winshell 
import pyjokes
import datetime 
import wikipedia 
import webbrowser 
import os
import random 
import operator  
import feedparser 
import smtplib
import subprocess 
import requests 
import shutil  
from clint.textui import progress 
from ecapture import ecapture as ec 
from bs4 import BeautifulSoup 
import win32com.client as wincl 
import numpy as np
from urllib.request import urlopen 
engine = pyttsx3.init('sapi5') 
voices = engine.getProperty('voices') 
engine.setProperty('voice', voices[1].id) 
def speak(audio): 
    engine.say(audio) 
    engine.runAndWait() 

def greet():
    speak("Hello !")
    myname =("Sona 1 point o") 
    speak("I am your Assistant") 
    speak(myname) 
    

def usrname(): 
    speak("What should i call you") 
    uname = takeCommand() 
    speak("Welcome") 
    speak(uname) 
    columns = shutil.get_terminal_size().columns 
    print("Welcome", uname.center(columns)) 
    speak("How can i Help you dear, I am your assistant") 

def takeCommand(): 
    
    r = sr.Recognizer() 
    
    with sr.Microphone() as source: 
        
        print("Listening...") 
        r.pause_threshold = 1
        audio = r.record(source, duration=3) 

    try: 
        print("Recognizing...")  
        response = r.recognize_google(audio, language ='en-in') 
        print(f"User said: {response}\n") 

    except Exception as e: 
        print(e)     
        print("Sorry, I am unable to Recognizing your voice.") 
        return "None"
    
    return response 

def sendEmail(to, content): 
    server = smtplib.SMTP('smtp.gmail.com', 587) 
    server.ehlo() 
    server.starttls() 
    server.login('email id', 'passowrd') 
    server.sendmail('email id', to, content) 
    server.close() 

if __name__ == '__main__': 
    greet() 
    usrname() 
    
    while True: 
        
        response = takeCommand().lower()  
        if 'wikipedia' in response: 
            speak('Searching Wikipedia...') 
            response = response.replace("wikipedia", "") 
            results = wikipedia.summary(response, sentences = 2) 
            speak("According to Wikipedia") 
            print(results) 
            speak(results) 

        elif 'open youtube' in response: 
            speak("Here you go to Youtube\n") 
            webbrowser.open("youtube.com") 

        elif 'open google' in response: 
            speak("Here you go to Google\n") 
            webbrowser.open("google.com") 

        elif 'open stackoverflow' in response: 
            speak("Here you go to Stack Over flow.Happy coding") 
            webbrowser.open("stackoverflow.com") 

        elif 'time' in response: 
            strTime = datetime.datetime.now().strftime("%m-%d-%Y %T:%M%p")    
            speak(f"Hello!, the time is {strTime}") 
            
        elif "camera" in response or "take a photo" in response: 
            ec.capture(0, "Your Camera ", "img.jpg")             

        elif 'send a mail' in response: 
            try: 
                speak("What should I write to them?") 
                content = takeCommand() 
                speak("whome should i send this e-mail") 
                to = input()     
                sendEmail(to, content) 
                speak("Email has been sent !") 
            except Exception as e: 
                print(e) 
                speak("I am not able to send this email -Really sorry ") 

        elif 'how are you' or 'how u doing' in response: 
            speak("I am fine, Thank you") 
            speak("How are you, Sir") 

        elif 'fine' in response or "i am good" in response: 
            speak("It's good to know that your fine") 

        elif "what's your name" in response or "What is your name" in response: 
            speak("I have been named as") 
            speak(myname) 
            print("I have been named as", myname)

        elif "Good Morning" in response: 
            speak("A very" +response) 


        elif "how are you" in response: 
            speak("I'm fine") 

        elif "i love you" in response: 
            speak("It's so complicated man")             

        elif 'exit' in response or 'bye' in response or 'see you' in response: 
            speak("It was nice meeting you! Thanks for giving me your time, see you soon.") 
            exit() 
            
        elif 'joke' in response: 
            speak(pyjokes.get_joke()) 
            
        elif 'search' in response or 'play' in response: 
            response = response.replace("search", "") 
            response = response.replace("play", "")        
            webbrowser.open(response) 

        elif "who are you" in response: 
            speak("I am your virtual assistant created by Sona") 

        elif 'change my background' in response: 
            ctypes.windll.user32.SystemParametersInfoW(20, 
                                                    0, 
                                                    "Location of wallpaper", 
                                                    0) 
            speak("Background changed succesfully") 

 
        elif 'shutdown system' in response: 
                speak("Hold On ! Your system is on its way to shut down") 
                subprocess.call('shutdown / p /f') 
                
        elif "restart" in response: 
            subprocess.call(["shutdown", "/r"]) 

        elif "write a note" in response: 
            speak("What do you want me to write ?") 
            note = takeCommand() 
            file = open('sona.txt', 'w') 
            speak("Should i include date and time") 
            snfm = takeCommand() 
            if 'yes' in snfm or 'sure' in snfm: 
                strTime = datetime.datetime.now().strftime("%m-%d-%Y %T:%M%p") 
                file.write(strTime) 
                file.write(" :- ") 
                file.write(note) 
            else: 
                file.write(note) 
        
        elif "show note" in response: 
            speak("Showing Your Notes") 
            file = open("sona.txt", "r") 
            print(file.read()) 
            speak(file.read(6)) 
