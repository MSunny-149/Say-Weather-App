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



    # Weather Code
    # pip install requests
    import requests
    speaker.Speak("Enter City Name: ")
    city = input("Enter City Name: ")
    url = f"https://api.weatherapi.com/v1/current.json?key=a3c511a0e6984172824182114242611&q={city}&aqi=yes"
    resp = requests.get(url)
    # print(resp.text)

    # Convert String resp.text (Response) data into Dictionary/JSON
    import json
    weatherData = json.loads(resp.text)
    temperature = weatherData["current"]["temp_c"]
    weatherText = f"The current weather of {city} is {temperature} degree celcius."
    print(weatherText)
    speaker.Speak(weatherText)

except:
    print("Please Turn On Your Internet and Try Again.")