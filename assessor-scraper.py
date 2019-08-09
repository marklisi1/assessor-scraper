import xlwt
import requests
from bs4 import BeautifulSoup

# initialize workbook and sheet
wb = xlwt.Workbook()
sheet = wb.add_sheet('Properties')
'''
# write headers
sheet.write(0,0,'Address')
sheet.write(0,1,'Hyperlink')
sheet.write(0,2,'Property Class')
sheet.write(0,3,'Municipal District')
sheet.write(0,4,'Assessment Area')
sheet.write(0,5,'2020 Total Value')
sheet.write(0,6,'2019 Total Value')
sheet.write(0,7,'2018 Total Value')
 # THERE ARE MORE, BUT LAZY. DO BY HAND
 '''


row = 1

list_of_links = True
list_of_addresses = False

if list_of_addresses:
    with open('addresses.txt') as f:
        for address in f:
            address = address.split()
            key = address[0] + '-' + address[1] + address[2] 
            url = 'http://qpublic9.qpublic.net/la_orleans_display.php?KEY=%s' % key

            sheet.write(row,1,url)
            html = requests.get(url)
            soup = BeautifulSoup(html.content, 'lxml')
            entries = soup.find_all('td')
            
            for i in range(len(entries)):
                if 'Location Address' in str(entries[i]):
                    sheet.write(row,0, entries[i+1].text)    
                if '2020' in str(entries[i]) and '*' in str(entries[i]):
                    price = ''.join(filter(lambda x: x.isdigit(), entries[i+3].text))
                    sheet.write(row,2, int(price))
                if 'Assessment Area' in str(entries[i]):
                    sheet.write(row,3, entries[i+1].text)
                if 'Municipal District' in str(entries[i]):
                    sheet.write(row,4, entries[i+1].text)
            print('property #%d written' % row)
            row += 1   

if list_of_links:
    with open('links.txt') as f:
        for link in f:
            link = link.strip()
            sheet.write(row,1,link)
            html = requests.get(link)
            soup = BeautifulSoup(html.content, 'lxml')
            entries = soup.find_all('td')

            for i in range(len(entries)):
                elem = entries[i].text.strip()
                if elem == 'Location Address':
                    sheet.write(row,0, entries[i+1].text)
                if elem == 'Property Class':
                    sheet.write(row,2, entries[i+1].text)
                if elem == 'Municipal District':
                    sheet.write(row,3, entries[i+1].text)  
                if elem == 'Assessment Area':
                    sheet.write(row,4, entries[i+1].text) 
                if elem == '*2020':
                    price = ''.join(filter(lambda x: x.isdigit(), entries[i+3].text))
                    sheet.write(row,5, int(price))
                    # Age and Disability Freeze
                    sheet.write(row,8, entries[i+9].text)
                    sheet.write(row,9, entries[i+10].text)
                if elem == '2019':
                    price = ''.join(filter(lambda x: x.isdigit(), entries[i+3].text))
                    sheet.write(row,6, int(price))
                    # Age and Disability Freeze
                    sheet.write(row,10, entries[i+9].text)
                    sheet.write(row,11, entries[i+10].text)
                if elem == '2018':
                    price = ''.join(filter(lambda x: x.isdigit(), entries[i+3].text))
                    sheet.write(row,7, int(price))
                    # Age and Disability Freeze
                    sheet.write(row,12, entries[i+9].text)
                    sheet.write(row,13, entries[i+10].text)

            print('property #%d written' % row)
            row += 1
            

wb.save('Properties.xls')