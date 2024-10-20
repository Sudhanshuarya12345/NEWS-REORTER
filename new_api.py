import  requests
import win32com.client as win
import json
from datetime import datetime, timedelta

speaker = win.Dispatch("SAPI.SpVoice")
api_key = 'afb6f6f5c60f429682e71794c3b809c2'

# find yesterday's date
today = datetime.now()
yesterday = today - timedelta(days=1)
date = yesterday.strftime("%Y-%m-%d")

def Give_news(url):
    response = requests.get(url)
    news_str = response.content.decode()
    parse_data = (json.loads(news_str))
    print(type(parse_data))
    
    i=0
    for data in parse_data['articles']:
        if i<5 : i+=1 
        else: break
        print("Author: ",data['author'])
        print("title: ",data['title'])
        print("description: ",data['description'])
        speaker.Speak(f'content: {data["content"]}')
        print()


while(1):
    # you can explore more and specify query such as cricket, weather, bihar weather, delhi weather, football, etc
    query = input("Enter your query: ")
    url = f'https://newsapi.org/v2/everything?q={query}&from={date}&to={date}&sortBy=popularity&apiKey={api_key}'

    Give_news(url)
    
    choice = input("Continue or not? \n('N' or 'n' for No, or press any other key to continue): ")
    if choice == 'N' or choice == 'n' : break

