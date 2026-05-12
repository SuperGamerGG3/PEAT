# Stock extension, created by Beffy

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
    pass