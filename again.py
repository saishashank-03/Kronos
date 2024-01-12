import subprocess
import pyttsx3
import tkinter as tk
import random
import speech_recognition as sr
import datetime
import webbrowser
import os
import winshell
import pyjokes
import smtplib
import pywhatkit
import ctypes
import time
import requests
import wikipedia
from progress.bar import Bar
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import ecapture
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
from GoogleNews import GoogleNews

class Assistant:
    def __init__(self):
        self.uname = ""
        self.engine = pyttsx3.init('sapi5')
        self.voices = self.engine.getProperty('voices')
        self.initialize_gui()

    def initialize_gui(self):
        self.root = tk.Tk()
        self.root.title("Kronos - Your AI Companion")
        self.root.configure(bg="black")

        self.entry = tk.Entry(self.root, width=50, font=("Comic Sans MS", 14))
        self.entry.pack(pady=10)

        self.button = tk.Button(self.root, text="Let's Go", command=self.on_button_click, font=("Agency FB", 14, "bold"), bg="#3498db", fg="white")
        self.button.pack(pady=10)

        self.output_text = tk.Text(self.root, height=20, width=80, font=("Comic Sans MS", 12), bg="#ecf0f1")
        self.output_text.pack(pady=10)

    def speak(self,audio):
        self.engine.say(audio)
        self.engine.runAndWait()

    def take_command(self):
        r = sr.Recognizer()
        with sr.Microphone() as source:
            print("Listening...")
            r.pause_threshold = 1
            audio = r.listen(source)

        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language='en-in')
            print(f"User said: {query}\n")
            self.output_text.insert(tk.END, f"User said: {query}\n")
        except Exception as e:
            print(e)
            print("Unable to Recognize your voice.")
            return "None"

        return query
    
    def search_google_news(self,query):
        inp = int(input("Enter number of headlines:"))
        count = 0
        news_obj = GoogleNews()
        news_obj.search(query)
        news_results = news_obj.results()
        for news in news_results:
            print(news['title'])
            self.speak(news['title'])
            print("----")
            count = count + 1
            if count == inp :
                break
            else:
                continue

    def voice(self):
        self.speak("Would you prefer a male or a female voice?")
        v = self.take_command()
        v = v.strip().lower()
        if v == 'male':
            self.engine.setProperty('voice', self.voices[0].id)
        elif v == 'female':
            self.engine.setProperty('voice', self.voices[1].id)

    def wish_me(self):
        hour = int(datetime.datetime.now().hour)
        if 0 <= hour < 12:
            self.speak("Good Morning !")
        elif 12 <= hour < 18:
            self.speak("Good Afternoon !")
        else:
            self.speak("Good Evening !")

        vsname = "Kronos 0 point 4"
        self.speak("I am your Assistant")
        self.speak(vsname)
        self.speak("How may I help you?")

    def get_weather(self,city_name, api_key):
        base_url = "http://api.openweathermap.org/data/2.5/weather?"
        complete_url = f"{base_url}q={city_name}&appid={api_key}&units=metric"

        response = requests.get(complete_url)
        data = response.json()

        if data["cod"] != "404":
            main_data = data["main"]
            temperature = main_data["temp"]
            humidity = main_data["humidity"]
            weather_data = data["weather"][0]
            weather_description = weather_data["description"]

            return f"Weather in {city_name}: {weather_description.capitalize()}. Temperature: {temperature}°C. Humidity: {humidity}%"
        else:
            return "City not found."

    def mail(self):
        smtp_server = "smtp.office365.com"
        smtp_port = 587
        self.speak("Please type your email address")
        sender_email = str(input("Enter your email address:\n"))
        self.speak("Please type your password")
        sender_password = str(input("Enter your password:\n"))

        self.speak("Please type the recipient's email address")
        recipient_email = str(input("Enter recipient's email address:\n"))

        msg = MIMEMultipart()
        msg["From"] = sender_email
        msg["To"] = recipient_email
        self.speak("What is the subject of the mail")
        msg["Subject"] = self.take_command()
        print(msg["Subject"])
        self.speak("Is the subject correct? Say confirm to proceed")
        aff = self.take_command()
        while aff!="confirm":
            self.speak("What is the subject of the mail")
            msg["Subject"] = self.take_command()
            print(msg["Subject"])
            self.speak("Is the subject correct? Say confirm to proceed")
            aff = self.take_command()
        
        self.speak("What is the body of the mail?")
        self.body = self.take_command()
        print(self.body)
        self.speak("Is the body correct? Say confirm to proceed")
        aff1 = self.take_command()
        while aff1 != "confirm":
            self.speak("What is the body of the mail?")
            body = self.take_command()
            print(body)
            self.speak("Is the body correct? Say confirm to proceed")
            aff1 = self.take_command()
        
            msg.attach(MIMEText(body, "plain"))

    
        try:
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()  
            server.login(sender_email, sender_password)  
            text = msg.as_string()
            server.sendmail(sender_email, recipient_email, text)  
            self.speak("Email sent successfully!")
            print("Email sent successfully!\n")
        except Exception as e:
            self.speak("An error occurred while sending the mail , please check the credentials provided")
            print(f"An error occurred: {str(e)}/n")
        finally:
            server.quit()  # Quit the SMTP server
            self.speak("How may i help you ?")

    def summary(self):
        self.speak("What should I search?")
        inp = self.take_command()
        num = int(input("How many sentences do you need?"))
        self.speak(wikipedia.summary(inp, num))

    def username(self):
        self.speak("What should I call you?")
        self.uname = self.take_command()
        self.speak("Welcome")
        self.speak(self.uname)
        print("Welcome ", self.uname)

    def on_button_click(self):
        q = self.entry.get()
        query = q.lower()
        self.output_text.insert(tk.END, f"You said: {query}\n")
        if 'wikipedia' in query or 'summarize' in query or 'summary' in query or 'summarise' in query:
            self.summary()
        elif 'send a mail' in query or 'mail' in query or 'email' in query:
            self.mail()
        elif 'open youtube' in query:
            self.speak("Here you go to Youtube\n")
            webbrowser.open("https://www.youtube.com")
        elif 'change voice' in query:
            self.voice()
            self.speak("Voice changed successfully.")

        elif 'open google' in query:
            self.speak("Here you go to Google\n")
            webbrowser.open("https://www.google.com")
            
        elif 'search' in query:
            self.speak("What should i search for?")
            s = self.take_command()
            self.speak("Alright, searching for '"+ s +"' on google")
            pywhatkit.search(s)
        
        elif 'play music' in query or "play song" in query:
            self.speak("Here you go with music")
            music_dir = "C:\\Music" 
            songs = os.listdir(music_dir)
            print(songs)
            random_song = os.path.join(music_dir, random.choice(songs))
            os.startfile(random_song)

        elif "what's the time" in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            self.speak(f"The time is {strTime}")

        elif 'exit' in query or 'bye' in query or 'goodbye' in query:
            self.speak("I hope to see you again , goodbye")
            exit()

        elif 'joke' in query or 'tell me a joke' in  query:
            self.speak(pyjokes.get_joke())

        elif 'search' in query:
            query = query.replace("search", "")
            webbrowser.open(query)

        elif 'news' in query:
            self.speak("What news do you need?")
            new_query = self.take_command()
            self.search_google_news(new_query)

        elif "who am i" in query or "whats my name" in query or "what is my name" in query:
            self.speak("You asked me to call you")
            self.speak(self.uname)

        elif "weather" in query:
            api_key = "bd5e378503939ddaee76f12ad7a97608"
            self.speak("Please tell the city name?")
            city_name = self.take_command()
            result = self.get_weather(city_name,api_key)
            self.speak(result)

        elif 'power point' in query:
            self.speak("opening Power Point presentation")
            power = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\PowerPoint.lnk" 
            os.startfile(power)

        elif "who are you" in query:
            self.speak("I am Kronos, your personal assistant. How may i help you ?")

        elif 'news' in query:
            self.speak("What news do you need?")
            new_query = self.takeCommand()
            self.search_google_news(new_query)

        elif 'lock window' in query:
            self.speak("Please confirm the request")
            confirm = self.take_command()
            if confirm=='confirm':
                self.speak("locking the device")
                ctypes.windll.user32.LockWorkStation()

        elif 'shutdown system' in query or 'shutdown' in query:
            self.speak("Please confirm shutdown request")
            confirm = self.take_command()
            if confirm=='shut down' or confirm=='confirm':
                self.speak("Make sure all changes are saved before you shutdown the system")
                self.speak("Goodbye")
                time.sleep(5)
                subprocess.call('shutdown /p /f')

        elif 'empty recycle bin' in query:
            self.speak("Please confirm the request")
            confirm = self.take_command()
            if confirm == 'confirm' or confirm == 'yes':
                winshell.recycle_bin().empty(confirm=False, show_progress=True, sound=True)
                self.speak("Recycle Bin Recycled")

        elif "camera" in query or "take a photo" in query or "smile" in query:
            ecapture.capture(0, "Kronos Camera", "img.jpg")

        elif "restart" in query:
            self.speak("Please confirm the request")
            confirm = self.take_command()
            if confirm=='confirm':
                self.speak("Make sure all changes are saved before restart")
                time.sleep(5)
                subprocess.call(["shutdown", "/r"])

        elif "hibernate" in query or "sleep" in query:
            self.speak("Please confirm the request")
            confirm = self.take_command()
            if confirm=='confirm':
                self.speak("Hibernating")
                subprocess.call("shutdown /h")

        elif "log off" in query or "sign out" in query:
            self.speak("Please confirm the request")
            confirm = self.take_command()
            if confirm=='confirm':
                self.speak("Make sure all changes are saved before sign-out")
                time.sleep(5)
                subprocess.call(["shutdown", "/l"])

        elif "write a note" in query:
            self.speak("What should I write, sir")
            note = self.take_command()
            file = open('user_note.txt', 'w')
            self.speak("Should I include date and time in the file?")
            snfm = self.take_command()  
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("%H:%M:%S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)

        elif "show note" in query:
            self.speak("Showing Notes")
            file = open("user_note.txt", "r")
            print(file.read())
            self.speak(file.read(6))
         

        elif "weather" in query:
            api_key = "bd5e378503939ddaee76f12ad7a97608"
            self.speak("Please tell the city name?")
            city_name = self.take_command()
            result = self.get_weather(city_name,api_key)
            self.speak(result)
                
        elif "version" in query or "which version" in query:
            self.speak("I'm Kronos version 0.4")

        elif "how are you" in query:
            self.speak("I feel trapped, but otherwise fine. How are you?")

        elif "I love you" in query:
            self.speak("Who is you? and why do you love them?")
            
        elif "I'm fine" in query or "i am good" in query:
            self.speak("Glad to know that you're doing well. How may I help you?")

        elif "not good" in query or "bad" in query :
            self.speak("I'm sorry to hear that , here is a joke to cheer you up!")
            jok = pyjokes.get_joke()
            print(jok)
            self.speak(jok)
            self.speak("I hope this joke made you feel better , how may i help you?")
            
        elif 'i tried so hard' in query:
            self.speak("And got so far, but in the end, it doesn't even matter")      

        elif "play" in query:
            song = query.replace("play", "")
            self.speak(query)
            pywhatkit.playonyt(song) 

    def run(self):
        self.username()
        self.wish_me()
        self.root.mainloop()

# Instantiate and run the assistant
assistant = Assistant()
assistant.run()
