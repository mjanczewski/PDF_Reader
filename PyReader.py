from pypdf import PdfReader
import re
import pandas as pd


reader = PdfReader("sandisk.pdf")
sandisk_szablon = pd.read_excel('sandisk.xlsx', dtype={'SO Line':str})
number_of_pages = len(reader.pages)
page = reader.pages[2]
text = page.extract_text()

# Regex ^[\d]{2}-[\d]{3}$
# line = "dfgdfg 51-111 "
# # reg = re.match('^[\d]{1}-[\d]{3}$', line)
# reg = re.search('[\d]{2}-[\d]{3}', line)
# print(reg)
# if reg: 
#     print(reg.group())

tablica_danych = []

dane = pd.DataFrame()


for page_number in range(number_of_pages-1):
    page = reader.pages[page_number]
    text = page.extract_text()
    # coo = text.find('COO')

    waga_wzorzec = re.compile(r'.[\d]\.[\d][\d][\d]')
    kraj_wzorzec = re.compile(r'[A-Z][A-Z]+(?=COO)')

    waga = waga_wzorzec.findall(text)
    kraj = kraj_wzorzec.findall(text)
    # print("Kraj: ", kraj)
    # waga = re.search(r'[\d][\d]\.[\d][\d][\d]', text)
    # if waga:
    #     print(waga)
    #     print(waga.group())
    # waga = text[1700:1722]

    # print(f'===== Strona numer {page_number}')
    # print(waga)
    
    # q = text.find("MYCOO")

    kraj_wzorzec = re.compile(r'[A-Z][A-Z]+(?=COO)')

    kraj = kraj_wzorzec.findall(text)

    q_wzorzec = re.compile(r'Quantity:')


    quan = q_wzorzec.finditer(text)


    suma_tablic = []

    iterator_pomocniczy = 0

    for w in quan:
        # print(f'Znak: "{w.group()}", miejsce: {w.start()+1}')
        # print(text[w.start()-6],text[w.start()-5],text[w.start()-4],text[w.start()-3],text[w.start()-2])
        ilosc = text[w.start()-7]+text[w.start()-6]+text[w.start()-5]+text[w.start()-4]+text[w.start()-3]+text[w.start()-2].rstrip()
        # print(ilosc)
        ilosc = ilosc.split(' ')
        # ======= TEST =========
        # print(waga)
        # waga = str(waga[1])
        # waga = waga.replace('.',',')
        # print(waga)
        # =========== END TEST ==========


# waga zamienic . na , 
# ilosć usunąć przecinek

        if len(kraj)>=2:
            if iterator_pomocniczy % 2==0:
                # print(kraj[0], waga[1], ilosc[1])
                # tablica_danych += ([kraj[0], ilosc[1], waga[1]])

                waga = str(waga[1])
                waga = waga.replace('.',',')
                tablica_danych.append([page_number+1, waga,ilosc[1],kraj[0], "Połącz kropki", f'Strona {page_number+1}'])
                iterator_pomocniczy +=1

                # ========= DZIAŁA START ===========
                # tablica_danych.append([waga[1],ilosc[1],kraj[0], "Połącz kropki", f'Strona {page_number+1}'])
                # iterator_pomocniczy +=1
                # ========= DZIAŁA END =============
            else:
                # print(kraj[1], waga[1], ilosc[1])
                # tablica_danych += ([kraj[1], ilosc[1], waga[1]])
                print(waga)
                waga = str(waga)
                waga = waga.replace('.',',')
                tablica_danych.append([f'{page_number+1}a',waga,ilosc[1],"", "Połącz kropki", f'Strona {page_number+1}'])
                iterator_pomocniczy +=1

                # ========= DZIAŁA START ===========
                # tablica_danych.append([waga[1],ilosc[1],"", "Połącz kropki", f'Strona {page_number+1}'])
                # iterator_pomocniczy +=1
                # ========= DZIAŁA END =============
        else:
            # tablica_danych += ([kraj[0], ilosc[1], waga[1]])

            # tablica_danych += [kraj[0], ilosc[1], waga[1]]

            waga = str(waga[1])
            waga = waga.replace('.',',')
            tablica_danych.append([page_number+1,waga,ilosc[1],kraj[0]])
            # ========= DZIAŁA START ===========
            # tablica_danych.append([waga[1],ilosc[1],kraj[0]])
            # ========= DZIAŁA END =============
            # print([kraj[0], waga[1], ilosc[1]])

        
    
# print(tablica_danych)

sandisk_szablon = sandisk_szablon.sort_values(by=['Delivery Number', 'SO #', 'SO Line'])
sandisk_szablon['strona'] = sandisk_szablon.iterrows

i = 1
for index, row in sandisk_szablon.iterrows():
    sandisk_szablon.loc[index, ['strona']] = i
    i += 1


df_dane = pd.DataFrame(tablica_danych)
# print(df_dane.columns)
df_nowa = pd.DataFrame()
df_nowa['Weight [kg]'] = df_dane[1]
df_nowa['COO Qty'] = df_dane[2]
df_nowa['COO'] = df_dane[3]
df_nowa['strona'] = df_dane[0]




connected = pd.merge(sandisk_szablon, df_nowa, left_on='strona', right_on='strona', how='right')
connected = connected.reindex(columns=['SO #', 'SO Line','Ship To',	'PO #',	'PO Date','Customer Part #','SKU','Ship Qty','Price','Currency','Ship Amt','Weight [kg]','COO Qty','COO','Ship Plant','Actual Ship Date','Estimated Delivery Date','Carrier','Tracking #','Incoterms','Ship To Address','Customer Name','Customer Address','Ship To #','Customer #','End Customer','Order Type','Customer Hier Name','Customer Hier #', 'Card Sub Type','Reporting Segment','Product Line','Delivery Number'
])


sandisk_szablon.to_excel('sandisk_mod.xlsx')
df_dane.to_excel('sandisk.xls')
connected.to_excel('sandisk_connected.xls', index=None)

# print(df_dane)



        # tablica_danych.append(tablica)


# for cos in range(len(tablica_danych)):

#     waga_total =tablica_danych[cos][2]
#     ilosc_total = tablica_danych[cos][1]
#     kraj_pochodzenia = tablica_danych[cos][0]
#     print(kraj_pochodzenia)
#     print(len(kraj_pochodzenia))
#     if len(kraj_pochodzenia)>=2:
#         for i in range(len(kraj_pochodzenia)):
#             print(kraj_pochodzenia[i])
#     print("="*100)




        
        

        




# ============================

# q = text.find("MYCOO")

# kraj_wzorzec = re.compile(r'[A-Z][A-Z]+(?=COO)')

# kraj = kraj_wzorzec.findall(text)

# q_wzorzec = re.compile(r'Quantity:')
# quan = q_wzorzec.finditer(text)
# for w in quan:
#     print(f'Znak: "{w.group()}", miejsce: {w.start()+1}')
#     print(text[w.start()-6],text[w.start()-5],text[w.start()-4],text[w.start()-3],text[w.start()-2])



