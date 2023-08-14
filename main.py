import psutil
import speech_recognition as sr
import win32com.client
import os
import webbrowser
import openai

# import these packages to ineract with the web
from selenium import webdriver
from selenium.webdriver.common.by import By

import subprocess

speaker = win32com.client.Dispatch("SAPI.SpVoice")

def say(query):
    speaker.Speak(query)
def takeCommand(user):
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source)
        try:
            # todo: language="ur" for URDU understanding & typing by Recognizer but not for speaking
            query = r.recognize_google(audio,language="en-pk")
            print(f"{user} : {query}")
            return query
        except Exception:
            return "Some Error Occured, Sorry."

def AI(prompt):

    openai.api_key = "API_KEY_HERE"

    prompt = prompt[:prompt.find('using GPT')].strip()
    output = f"GPT response for prompt: {prompt} \n ********************** \n\n"

    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo-16k-0613",
        messages=[
            {
                "role": "user",
                "content": prompt
            }
        ],
        temperature=1,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    try:
        print(response["choices"][0]["message"]["content"])
    except Exception:
        return Exception;

    output += response["choices"][0]["message"]["content"]

    if not os.path.exists("Openai"):
        os.mkdir("Openai")
    with open(f"Openai/{prompt}.txt","w") as f:
        f.write(output)

def chat_AI(prompt):

    # TODO: write your API key for openai here
    openai.api_key = "API_KEY_HERE"

    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo-16k-0613",
        messages=[
            {
                "role": "user",
                "content": prompt
            }
        ],
        temperature=1,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    # Exception is used as there might be a response with zero choices so in that case Exception will be thrown and caught here
    try:
        print(f"SOAL: {response['choices'][0]['message']['content']}\n ")
        say(response["choices"][0]["message"]["content"])
    except Exception:
        return Exception


def end_Processes(processName):

    # psutil is a safe way to close program as it does not end processes forcefully
    for process in psutil.process_iter(attrs=['pid', 'name']):
        if processName in process.info['name']:
            psutil.Process(process.info['pid']).terminate()
            print(f"closed {processName}")


def initialize_driver():
    # opening the youtube in chromedriver
    # driver = webdriver.Chrome(executable_path="C:\\Users\\Shakaib\\Desktop\\chromedriver.exe")
    # I haven't specified the driver path as I had set it in environment variables "PATH"
    global driver
    driver = webdriver.Chrome()

def searchYoutube(search_query):
    # find the search bar using selenium find_element function
    driver.find_element(By.NAME, "search_query").send_keys("")
    driver.find_element(By.NAME, "search_query").send_keys(search_query)

    # clicking on the search button
    driver.find_element(By.CSS_SELECTOR, "#search-icon-legacy.ytd-searchbox").click()
def automateYoutube():

    say("Tell me, what you want to search on youtube?")
    search_query = takeCommand()

    while "exit" not in search_query:

        if "search" in search_query:
            searchYoutube(search_query)
        if "play" in search_query:
            # Play the video
            play_button = driver.find_element(By.CSS_SELECTOR, ".ytp-play-button")
            play_button.click()

        if "stop" in search_query:
            # Pause the video
            pause_button = driver.find_element(By.CSS_SELECTOR, ".ytp-play-button")
            pause_button.click()

        if "play next" in search_query:
            # Play the next video (click on the next button)
            next_button = driver.find_element(By.CSS_SELECTOR, ".ytp-next-button")
            next_button.click()
        '''
        if "play previous" in search_query:
            # Play the previous video (click on the previous button)
            previous_button = driver.find_element(By.CSS_SELECTOR, ".ytp-prev-button")
            previous_button.click()
        '''
        print("Listening For Youtube")
        search_query = takeCommand()

if __name__ == '__main__':

    print('PyCharm')
    say("Hello sir I'm Soal")
    say("Kindly, Tell me your name?")

    user = input("Name: ")
    say(f"Welcome! {user}. How can I help you today?")

    while True:
        pass_itteration = False

        print("Listening........")
        query = takeCommand(user)

        sites = [ ["wikipedia","https://www.wikipedia.com"],["Git Hub","https://github.com/S-H-A-B-B-Y?tab=repositories"]
               , ["netflix","https://www.netflix.com/browse"], ["google","https://www.google.com"]
               , ["GPT","https://chat.openai.com"], ["movies","https://theflixer.tv/movie"]
               , ["Anime","https://ww4.gogoanime2.org"]]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                webbrowser.open(site[1])
                say(f"Opening {site[0]} for you")
                pass_itteration = True

    # Used list of lists here first parameter is the name of the game, second parameter is the path of the game, third one is the name of the game executable file used to close it
        games = [
            ["Epic Games", "C:\Program Files (x86)\Epic Games\Launcher\Portal\Binaries\Win32\EpicGamesLauncher.exe","EpicGamesLauncher.exe"]
            , ["Sacro", "D:\Sekiro - Shadows Die Twice\sekiro.exe","sekiro.exe"],
            ["Last of Us", "D:\The Last of Us - Part I\launcher.exe","launcher.exe"]
            , ["Uncharted", "D:\\UNCHARTED - Legacy of Thieves Collection\Launcher.exe","Launcher.exe"]
            , ["Cricket", "D:\Cricket 22\cricket22.exe","cricket22.exe"]]
        for game in games:
            if f"open {game[0]}".lower() in query.lower():
                os.startfile(game[1])
                say(f"Opening {game[0]} for you")
                pass_itteration = True
            if f"close {game[0]}".lower() in query.lower():
                end_Processes(game[2])
                pass_itteration = True

        if pass_itteration:
            continue

        if "Using GPT".lower() in query.lower():
            AI(query)

        elif "open youtube" in query.lower():
            say(f"Opening youtube for you")
            initialize_driver()
            url = "https://www.youtube.com/"
            driver.get(url)
            automateYoutube()

    # Its just a template to show that songs can also be played using OS .startfile() and you can implement your own songs using this template
        elif "open poison" in query:
            musicPath="D:\songs\Poison.mp4"
            os.startfile(musicPath)

    # This particular process would not run if you haven't started your script using administrator privileges like "run as administrator"
    # because it is a system process with sensitive privileges
        elif "open Task Manager".lower() in query.lower():
            subprocess.Popen(["taskmgr.exe"])
        elif "close Task Manager".lower() in query.lower():
            end_Processes('Taskmgr.exe')

        else:
            chat_AI(query)


