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


    # Weather API Code (Old API)
    url = f"https://api.weatherapi.com/v1/current.json?key=a3c511a0e6984172824182114242611&q={city}&aqi=yes"


    # Internet Code
    # pip install requests
    import requests
    resp = requests.get(url)
    respData = resp.text # Storing Full Weather Response (String Data)
    

    # Converting String respData (Response) data into Dictionary/JSON
    import json
    weatherData = json.loads(respData)
    temperature = weatherData["current"]["temp_c"] # Storing Only Temperature
    weatherText = f"The current weather of {city} is {temperature} degree celcius."
    print(weatherText)
    speaker.Speak(weatherText)


except:
    internetOffMssg = "Please Turn On Your Internet and Try Again."
    speaker.Speak(internetOffMssg)
    print(internetOffMssg)