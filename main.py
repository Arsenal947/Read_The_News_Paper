from win32com.client import Dispatch
import requests
import json
from datetime import date


def speak(str):
    speak = Dispatch("SAPI.spVoice")
    speak.Speak(str)


url = "https://newsapi.org/v2/top-headlines?country=in&apiKey=11abe30c21c7456d86aa27b2aeeb76e8"

if __name__ == "__main__":
    speak(f"news for {date.today()}")
    news = requests.get(url).text
    news_dict = json.loads(news)
    arti = news_dict['articles']

    for titles in arti:
        aritcles = titles['title']
        speak(aritcles)
        if len(aritcles) - 1:
            speak("And the next News is")

    speak("Thank You for listening")
