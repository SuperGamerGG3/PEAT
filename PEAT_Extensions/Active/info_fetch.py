"""
PEAT EXTENSION
author = Beffy
name = Internet Info Fetch
filename = info_fetch
version = 1.0

requirements:
- None
"""

import urllib
import xml.etree.ElementTree as ET

EXT_NAMESPACE = "inf"

help_dict = {
    "weather": "Fetches weather",
    "news": "Gets news from specified short site"
}

def load_extension():
    peat.register_command(EXT_NAMESPACE, "weather", cmd_weather) # type: ignore
    peat.register_command(EXT_NAMESPACE, "news", cmd_news) # type: ignore
    
    peat.register_help(EXT_NAMESPACE, help_dict) # type: ignore

def fetch_weather(place):
    safe_place = urllib.parse.quote(place)
    url = f"https://wttr.in/{safe_place}?format=4"

    try:
        response = urllib.request.urlopen(url)
    except Exception as e:
        peat.log(f"Couldn't get weather: {e}") # type: ignore
        return f"Exception: {e}"

    return response.read().decode("utf-8").strip()

def cmd_weather(a1, a2, title):
    try:
        location_raw = a1.strip() if a1 else ""

        data = fetch_weather(a1)

        print("\n--- Weather Report ---")
        peat.voice_print(f"{data}") # type: ignore
        print("----------------------")

        peat.log(f"Weather fetched: {location_raw}") # type: ignore

    except Exception as e:
        peat.voice_print(f"Failed to fetch weather: {e}") # type: ignore
        peat.log(f"Weather error: {e}") # type: ignore

def cmd_news(a1, a2, title):
    items = a2
    if a1 == "help":
        print("News sources:")
        print("- hack: Hacker News")
        print("- bbc: BBC News")
        print("- abc: ABC News")
        print("- guard: The Guardian US")
        print("- nbc: NBC News Politics")
        print("- fox: Fox News Politics")
        print("News sources displayed.")
        return
    
    if a1 == "hack":
        url = "https://news.ycombinator.com/rss"
        a1 = "Hacker News"
    elif a1 == "bbc":
        url = "http://feeds.bbci.co.uk/news/rss.xml"
        a1 = "BBC News"
    elif a1 == "abc":
        url = "https://abcnews.com/abcnews/usheadlines"
        a1 = "ABC News"
    elif a1 == "guard":
        url = "https://www.theguardian.com/us-news/rss"
        a1 = "The Guardian"
    elif a1 == "nbc":
        url = "https://feeds.nbcnews.com/nbcnews/public/politics"
        a1 = "NBC News"
    elif a1 == "fox":
        url = "https://moxie.foxnews.com/google-publisher/politics.xml"
        a1 = "FOX News"
    elif a1 == "":
        print("Expected quoted string for news source, but got nothing.")
        return
    else:
        peat.voice_print("Unknown news source.") # type: ignore
        return

    if items != "":
        try:
            items_to_fetch = int(items)
        except ValueError:
            items_to_fetch = peat.news_items_per_request # type: ignore
    else:
        items_to_fetch = peat.news_items_per_request # type: ignore

    try:
        response = urllib.request.urlopen(url)
        data = response.read()

        root = ET.fromstring(data)

        print(f"--- Latest News from {a1} ---")

        count = 0
        for item in root.iter("item"):
            title_tag = item.find("title")
            if title_tag is None:
                continue
            title = title_tag.text 
            print(f"- {title}")

            count += 1
            if count >= items_to_fetch:  # Limit to 7 news items
                break

        print("-------------------")
        peat.log("Fetched news successfully") # type: ignore
        peat.voice_print("News fetched successfully.") # type: ignore

    except Exception as e:
        peat.voice_print(f"Failed to fetch news: {e}") # type: ignore
        peat.log(f"News fetch error: {e}") # type: ignore