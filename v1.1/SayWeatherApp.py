try:
    # Speaking Code
    # pip install pypiwin32
    import win32com.client
    speaker = win32com.client.Dispatch("SAPI.SpVoice") 

    # while True:
    #     text = input("Enter the word you want to speak: ")
    #     if (text == "exit" or text == "Exit"):
    #         break
    #     speaker.Speak(text)


    # Getting City from User
    speaker.Speak("Enter City Name: ")
    city = input("Enter City Name: ")


    # Getting Latitude & Longitude from City Code
    from geopy.geocoders import Nominatim
    # Initialize Nominatim API
    geolocator = Nominatim(user_agent="MyApp")
    location = geolocator.geocode(city)
    locLati = location.latitude
    locLong = location.longitude

    
    # Weather API Code (New API)
    url = f"https://api.open-meteo.com/v1/forecast?latitude={locLati}&longitude={locLong}&current=temperature_2m"


    # Internet Code
    # pip install requests
    import requests
    resp = requests.get(url)
    respData = resp.text # Storing Full Weather Response (String Data)


    # Convert String respData (Response) data into Dictionary/JSON
    import json
    weatherData = json.loads(respData)
    temperature = weatherData["current"]["temperature_2m"] # Storing Only Temperature
    weatherText = f"The current weather of {city} is {temperature} degree celcius."
    print(weatherText)
    speaker.Speak(weatherText)


except:
    internetOffMssg = "Please Turn On Your Internet and Try Again."
    speaker.Speak(internetOffMssg)
    print(internetOffMssg)