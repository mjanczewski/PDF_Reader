from pypdf import PdfReader
import re
import pandas as pd


reader = PdfReader("sandisk.pdf")
sandisk_szablon = pd.read_excel('sandisk.xlsx', dtype={'SO Line':str})
number_of_pages = len(reader.pages)
page = reader.pages[0]
text = page.extract_text()
tablica_danych = []
dane = pd.DataFrame()


for page_number in range(number_of_pages-1):
    page = reader.pages[page_number]
    text = page.extract_text()

    waga_wzorzec = re.compile(r'..[\d]\.[\d][\d][\d]')
    kraj_wzorzec = re.compile(r'[A-Z][A-Z]+(?=COO)')

    waga = waga_wzorzec.findall(text)
    kraj = kraj_wzorzec.findall(text)
    kraj_wzorzec = re.compile(r'[A-Z][A-Z]+(?=COO)')
    kraj = kraj_wzorzec.findall(text)
    q_wzorzec = re.compile(r'Quantity:')
    quan = q_wzorzec.finditer(text)
    suma_tablic = []
    iterator_pomocniczy = 0


    for w in quan:

        ilosc = text[w.start()-8]+text[w.start()-7]+text[w.start()-6]+text[w.start()-5]+text[w.start()-4]+text[w.start()-3]+text[w.start()-2].rstrip()
        print(ilosc)
        ilosc = ilosc.split(' ')
         
        if len(kraj)>=2:
            if iterator_pomocniczy % 2==0:
                waga = str(waga[1])
                waga = waga.replace('.',',')
                tablica_danych.append([page_number+1, waga,ilosc[1],kraj[0], "Połącz kropki", f'Strona {page_number+1}'])
                iterator_pomocniczy +=1
            else:
                waga = str(waga)
                waga = waga.replace('.',',')
                tablica_danych.append([f'{page_number+1}a', '',ilosc[1],kraj[1],"", "Połącz kropki", f'Strona {page_number+1}'])
                iterator_pomocniczy +=1
        else:
            
            if len(ilosc)>=2:
                waga = str(waga[1])
                waga = waga.replace('.',',')
                tablica_danych.append([page_number+1,waga,ilosc[1],kraj[0]])
            else:
                waga = str(waga[1])
                waga = waga.replace('.',',')
                tablica_danych.append([page_number+1,waga,ilosc[0],kraj[0]])


sandisk_szablon = sandisk_szablon.sort_values(by=['Delivery Number', 'SO #', 'SO Line'])


i = 1
for index, row in sandisk_szablon.iterrows():
    sandisk_szablon.loc[index, ['strona']] = i
    i += 1


df_dane = pd.DataFrame(tablica_danych)
df_nowa = pd.DataFrame()
df_nowa['Weight [kg]'] = df_dane[1]
df_nowa['COO Qty'] = df_dane[2]
df_nowa['COO'] = df_dane[3]
df_nowa['strona'] = df_dane[0]


connected = pd.merge(sandisk_szablon, df_nowa, left_on='strona', right_on='strona', how='right')
connected = connected.reindex(columns=['SO #', 'SO Line','Ship To',	'PO #',	'PO Date','Customer Part #','SKU','Ship Qty','Price','Currency','Ship Amt','Weight [kg]','COO Qty','COO','Ship Plant','Actual Ship Date','Estimated Delivery Date','Carrier','Tracking #','Incoterms','Ship To Address','Customer Name','Customer Address','Ship To #','Customer #','End Customer','Order Type','Customer Hier Name','Customer Hier #', 'Card Sub Type','Reporting Segment','Product Line','Delivery Number'
])

connected.to_excel('sandisk_connected.xls', index=None)