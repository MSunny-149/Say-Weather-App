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
    # AQI API Code
    url_AQI = f"https://air-quality-api.open-meteo.com/v1/air-quality?latitude=23.0258&longitude=72.5873&current=us_aqi"


    # Internet Code - Getting Response Data as String
    # pip install requests
    import requests
    # Weather
    resp = requests.get(url) # Response Code
    respData = resp.text # Storing Full Weather Response (String Data)
    # AQI
    resp_AQI = requests.get(url_AQI) # Response Code
    resp_AQI_Data = resp_AQI.text # Storing AQI Response (String Data)


    # Convert String Response Data (Text) into Dictionary/JSON
    import json
    # Weather
    weatherData = json.loads(respData)
    temperature = weatherData["current"]["temperature_2m"] # Storing Only Temperature
    # AQI
    AQIData = json.loads(resp_AQI_Data)
    AQI = AQIData["current"]["us_aqi"]
    if (AQI >= 0 and AQI <= 19):
        scaleOfAQI = "Excellent"
    elif (AQI >= 20 and AQI <= 49):
        scaleOfAQI = "Fair"
    elif (AQI >= 50 and AQI <= 99):
        scaleOfAQI = "Poor"
    elif (AQI >= 100 and AQI <= 149):
        scaleOfAQI = "Unhealthy"
    elif (AQI >= 150 and AQI <= 249):
        scaleOfAQI = "Very Unhealthy"
    elif (AQI >= 250):
        scaleOfAQI = "Dangerous"


    # Final Message to Speak
    finalMessage = f"The current weather of {city} is {temperature} degree celcius with an AQI of {AQI} which is {scaleOfAQI}."
    print(finalMessage)
    speaker.Speak(finalMessage)


except:
    internetOffMssg = "Please Turn On Your Internet and Try Again."
    speaker.Speak(internetOffMssg)
    print(internetOffMssg)