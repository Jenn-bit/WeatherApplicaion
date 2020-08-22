import xlwings as xl
import requests, json


wb= xl.Book(r"WeatherApp.xlsx")
sh = wb.sheets['Weather']


def find_weather(city_name):

    api_key = "afd3ae5dbabb693461e327ec22c7ec0e"

    base_url = "http://api.openweathermap.org/data/2.5/weather?"
    complete_url = base_url + "appid=" + api_key + "&q=" + city_name
    response = requests.get(complete_url)
    x = response.json()
    if x["cod"] != "404":
        y = x["main"]
        current_temperature = y["temp"]
        current_humidiy = y["humidity"]
        result = (current_temperature, current_humidiy)

    else:
        result = (0, 0)
    return result





def update_weather():

    cells_down = sh.range('A1').end('down').row

    while cells_down>2:
        s = str(cells_down)
        update = sh.range('E'+s).value
        if update == 1:
            city_name=sh.range('A'+s).value
            weather = find_weather(city_name)
            temp = weather[0]
            humidity = weather[1]
            if temp == 0 and humidity == 0:
                sh.range('B' + s).value = '-'
                sh.range('C' + s).value = '-'
            else:
                unit = sh.range('D'+s).value
                if unit == 'c' or unit == 'C':
                    temp = temp - 273.15
                    sh.range('B' + s).value = temp
                else:
                    temp = temp*(9/5) - 459.67
                    sh.range('B' + s).value = temp

                sh.range('C'+s).value = humidity

        cells_down=cells_down-1


while 'true':
    update_weather()

