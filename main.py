import psutil
import speech_recognition as sr
import win32com.client
import os
import webbrowser

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.common.by import By

import subprocess

speaker = win32com.client.Dispatch("SAPI.SpVoice")

def say(query):
    speaker.Speak(query)
def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio,language="en-pk") #language="ur" for URDU typing but not for speaking
            print(f"User input: {query}")
            return query
        except Exception:
            return "Some Error Occured, Sorry."

def end_Processes(processName):
    for process in psutil.process_iter(attrs=['pid', 'name']):
        if processName in process.info['name']:
            psutil.Process(process.info['pid']).terminate()
            print(f"closed {processName}")
def initialize_driver():
    # opening the youtube in chromedriver
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

        if "pause" in search_query:
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
    while True:
        print("Listening........")
        query = takeCommand()
        sites = [ ["wikipedia","https://www.wikipedia.com"]
               , ["netflix","https://www.netflix.com/browse"], ["google","https://www.google.com"]
               , ["GPT","https://chat.openai.com"], ["movies","https://theflixer.tv/movie"]
               , ["Anime","https://ww4.gogoanime2.org"]]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                webbrowser.open(site[1])
                say(f"Opening {site[0]} for you")
        if "open youtube" in query.lower():
            say(f"Opening youtube for you")
            initialize_driver()
            url = "https://www.youtube.com/"
            driver.get(url)
            automateYoutube()
        if "open poison" in query:
            musicPath="D:\songs\Poison.mp4"
            os.startfile(musicPath)
        games = [
            ["Epic Games", "C:\Program Files (x86)\Epic Games\Launcher\Portal\Binaries\Win32\EpicGamesLauncher.exe","EpicGamesLauncher.exe"]
            , ["Sacro", "D:\Sekiro - Shadows Die Twice\sekiro.exe"],
            ["Last of Us", "D:\The Last of Us - Part I\launcher.exe"]
            , ["Uncharted", "D:\\UNCHARTED - Legacy of Thieves Collection\Launcher.exe"]
            , ["Cricket", "D:\Cricket 22\cricket22.exe"]]
        for game in games:
            if f"open {game[0]}".lower() in query.lower():
                os.startfile(game[1])
            if f"close {game[0]}".lower() in query.lower():
                end_Processes(game[2])
        if "open Task Manager".lower() in query.lower():
            subprocess.Popen(["taskmgr.exe"])
        if "close Task Manager".lower() in query.lower():
            end_Processes('Taskmgr')


