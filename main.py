from bs4 import BeautifulSoup
import requests
import pandas as pd
import xlsxwriter


#Formatting
web = ['https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=1',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=2',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=3',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=4',
        'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=5',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=6',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=7',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=8',
        'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=9',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=10',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=11',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=12',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=13',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=14',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=15',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=16',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=17',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=18',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=19',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=20',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=21',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=22',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=23',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=24',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=25',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=26',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=27',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=28',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=29',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=30',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=31',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=32',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=33',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=34',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=35',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=36',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=37',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=38',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=39',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=40',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=41',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=42',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=43',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=44',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=45',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=46',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=47',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=48',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=49',
       'https://www.century21global.com/for-sale-residential/Hong-Kong?sort=dateDesc&pageNo=50',
       ]

Location = []
Size = []
Price_HKD = []
Price_USD = []

Data = {"Size": [],"Location":[],"Price_HKD" : [],"Price_USD" : []}

Size = []
Location = []
Price_HKD = []
Price_USD = []


for pg in web:
    webpage_response = requests.get(pg)
    webpage = webpage_response.content
    soup = BeautifulSoup(webpage,'html.parser')
    sizes = soup.find_all(attrs={'class': 'size'})

    for s in sizes:
        for i in s:
            size = s.get_text()
            if (s.index(i) + 1)/2 != 1:
                Size.append(str.replace(size,"\r\n                                    ",""))

    prices_HKD = soup.find_all(attrs={'class':"search-result-label-primary price-native"})
    for p in prices_HKD:
        for i in p:
            price = p.get_text()
            if (p.index(i) + 1)/2 != 1:
                Price_HKD.append(price)

    prices_USD = soup.find_all(attrs={'class':"search-result-label-secondary price-user"})
    for p in prices_USD:
        for i in p :
            price = p.get_text()
            if (p.index(i) + 3) % 3 == 0:
                Price_USD.append(price)

    location = soup.find_all(attrs={'class':"search-result-label"})
    for l in location:
        for i in l:
            lo = l.get_text()
            if 'Hong Kong' in lo:
                Location.append(str.replace(lo,"\r\n                        ",""))






Data['Size'] = Size
Data['Location'] = Location
Data['Price_HKD'] = Price_HKD
Data['Price_USD'] = Price_USD

Propertydf = pd.DataFrame.from_dict(Data)

print(Propertydf)
writer_obj = pd.ExcelWriter('property.xlsx',
                            engine='xlsxwriter')


Propertydf.to_excel(writer_obj, sheet_name='Sheet')

writer_obj.save()