import requests
import xlsxwriter
import xmltodict

# get data from efis url
url = "https://efis.mcx.ru/ui/default/apps/opendata/%D0%A1%D0%BF%D1%80%D0%B0%D0%B2%D0%BE%D1%87%D0%BD%D0%B8%D0%BA_%D0%BA%D1%83%D0%BB%D1%8C%D1%82%D1%83%D1%80.xml"
response = requests.get(url)
data = response.content
d = xmltodict.parse(data)

#local log file
my_file = open("C:/Users/kharlamov/Desktop/Работа/Иркутск/parcer/cult.txt", "w")

#write data to xlsx
workbook = xlsxwriter.Workbook('C:/Users/kharlamov/Desktop/Работа/Иркутск/parcer/cult.xlsx')
worksheet = workbook.add_worksheet()

print(d['Workbook']['ss:Worksheet']['Table']['Row'][12]['Cell'][1]['Data']['#text'])

count = 1

for rows in d['Workbook']['ss:Worksheet']['Table']['Row']:
    try:
        print(rows['Cell'][0]['Data']['#text'] + " - " + rows['Cell'][1]['Data']['#text'] + "\n")
        my_file.write(rows['Cell'][0]['Data']['#text'] + " - " + rows['Cell'][1]['Data']['#text'] + "\n")
        worksheet.write(count, 0, rows['Cell'][0]['Data']['#text'])
        worksheet.write(count, 1, rows['Cell'][1]['Data']['#text'])
        count += 1

    except Exception:
        print("Bruh...")

workbook.close()
my_file.close()

