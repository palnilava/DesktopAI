import webbrowser
import speech_recognition as sr
import os,subprocess
import random
import win32com.client
import datetime
import openai
#from config import apikey

chatstr=""
apikey = os.getenv("OPENAI_API_KEY")
def chat(query,chatstr):
    print(chatstr)
    openai.api_key = apikey
    chatstr += f" Nilava : {query}\n AI: "
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=chatstr,
        temperature=1,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )

    #print(response["choices"][0]["text"])
    print(chatstr + f"{response['choices'][0]['text']}" )
    s.Speak(response["choices"][0]["text"])
    chatstr += f" {response['choices'][0]['text']} "
    return response["choices"][0]["text"]


def ai(prompt):
    openai.api_key = apikey
    text = f" Prompt -{prompt}\n\n"
    response = openai.Completion.create(
      model="text-davinci-003",
      prompt=prompt,
      temperature=1,
      max_tokens=256,
      top_p=1,
      frequency_penalty=0,
      presence_penalty=0
    )

    print(response["choices"][0]["text"])
    text += response["choices"][0]["text"]

    if not os.path.exists("openai"):
        os.mkdir("openai")

    with open(f"openai/{prompt.split('AI',1)[1]}.txt" , "w") as f:
        f.write(text)

s = win32com.client.Dispatch("SAPI.SpVoice")
def takecommand():
     r = sr.Recognizer()
     with sr.Microphone() as source:
         r.pause_threshold = 1
         audio = r.listen(source)
         try:
             print("Recognizing....")
             query = r.recognize_google(audio,language="en-in")
             print(f"User said {query}")
             return query
         except Exception as e:
             return  "Some Error Occured"

if __name__ == '__main__':
    print("Starting....")
    s.Speak("Hello say Something to your AI bot")
    while True:
        print("Listening....")
        query = takecommand()
        # s.Speak(query)
        sites= [["youtube","https://www.youtube.com/"], ["google","https://www.google.com/"],["facebook","https://www.facebook.com/"]]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                s.Speak(f" Opening {site[0]} sir..")
                webbrowser.open(site[1])

        if "open music" in query:
           musicPath = "C:/Users/asisp/OneDrive/Documents/Music"
           files = os.listdir(musicPath)
           d = random.choice(files)
           os.startfile(f"C:/Users/asisp/OneDrive/Documents/Music/{d}")

        if "the time" in query:
            now = datetime.datetime.now()
            s.Speak(f" Sir the time now is {now.hour} hours, {now.minute} minutes,{now.second} second ")

        #AUMID - Application User Model ID'
        apps = [["firefox", "start explorer shell:appsfolder/308046B0AF4A39CB"], ["clock", "start explorer shell:appsfolder/Microsoft.WindowsAlarms_8wekyb3d8bbwe!App"],
                 ["calculator", "start explorer shell:appsfolder/Microsoft.WindowsCalculator_8wekyb3d8bbwe!App"],["camera", "start explorer shell:appsfolder/Microsoft.WindowsCamera_8wekyb3d8bbwe!App"],
                ["calendar", "start explorer shell:appsfolder/microsoft.windowscommunicationsapps_8wekyb3d8bbwe!microsoft.win..."],["weather", "start explorer shell:appsfolder/Microsoft.BingWeather_8wekyb3d8bbwe!App"]]
        for app in apps:
            if f"Open {app[0]}".lower() in query.lower():
                s.Speak(f" Opening {app[0]} sir..")
                os.system(app[1])

        if "using ai".lower() in query.lower():
            ai(prompt=query)

        if "stop".lower() in query.lower():
            s.Speak("Exiting from your AI bot")
            break

        else:
            chat(query,chatstr)