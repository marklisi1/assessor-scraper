import xlwt
import requests
from bs4 import BeautifulSoup

wb = xlwt.Workbook()
sheet = wb.add_sheet('Properties')

#write headers
sheet.write(0,0,'Address')
sheet.write(0,1,'Hyperlink')
sheet.write(0,2,'Total Value (2020)')
sheet.write(0,3,'Assessment Area')
sheet.write(0,4,'Municipal District')


url = 'http://qpublic9.qpublic.net/la_orleans_display.php?KEY=2201-JEFFERSONAV'
sheet.write(1,1,url)

html = requests.get(url)

soup = BeautifulSoup(html.content, 'lxml')
print(soup.prettify())
entries = soup.find_all('td')
for i in range(len(entries)):

    if 'Location Address' in str(entries[i]):
        sheet.write(1,0, entries[i+1].text)    

    if '2020' in str(entries[i]) and '*' in str(entries[i]):
        price = ''.join(filter(lambda x: x.isdigit(), entries[i+3].text))
        sheet.write(1,2, int(price))

    if 'Assessment Area' in str(entries[i]):
        sheet.write(1,3, entries[i+1].text)
    
    if 'Municipal District' in str(entries[i]):
        sheet.write(1,4, entries[i+1].text)

    tables = soup.find_all('table')


wb.save('Properties.xls')


