import spacy
import huspacy
from openpyxl import load_workbook
import time
#huspacy.download
DATA_PATH = ".\\ossz.xlsx"
nlp = spacy.load("hu_core_news_lg")
book = load_workbook(DATA_PATH,data_only=True)
sheet = book["Egybe"]

tanarnevek = ['Duret József','Ágoston György', 'Albert Sándor', 'Almási Tibor', 'Anderle Ádám', 'Bakos Ferenc', 'Balázs Mihály', 'Bálint Alajos', 'Bálint Sándor', 'Balogh Tibor', 'Banner János', 'Bánréti Zoltán', 'Baranyai Erzsébet', 'Baranyai Zoltán', 'Bárczi Géza', 'Barna Gábor', 'Baróti Dezső', 'Baróti Tiborné 1. Gaál Márta', 'Bartók György', 'Bassola Péter', 'Belényi Gyula', 'Bellon Tibor', 'Benedek Nándor', 'Bényiné 1. Farkas Mária', 'Bérezik Árpád', 'Bernáth Árpád', 'Berta Árpád', 'Birkás Géza', 'Boda István', 'Bognár Cecil Pál', 'Bollobás Enikő', 'Buday Árpád', 'Czachesz Erzsébet Cs.', 'Csapó Benő', 'Csefkó Gyula', 'Csejtei Dezső','Csengery János', 'Cseresnyési László', 'Csetri Lajos', 'Csukás István', 'Csúri Károly', 'Deér József', 'Deme László', 'Dézsi Lajos', 'Domokos Péter', 'Dornbach Mária', 'Duró Lajos', 'Eördögh István', 'Eperjessy Kálmán', 'Erdélyi Gyula', 'Erdélyi László', 'Erdélyi László Gyula', 'Erős Ferenc', 'Fabiny Tibor', 'Fábricz Károly', 'Farkas Mária', ' Bényiné', 'Fazekas Erzsébet', ' Gerő Ernőné', 'Fejér Ádám', 'Fejes Katalin', ' B.', 'Feleky Gábor', 'Felvinczi Takáts Zoltán', 'Fenyvesi István', 'Ferenczi Imre', 'Fodor István', 'Fogarasi Miklós', 'Fógel József', 'Fónagy Iván', 'Forgács Tamás', 'Förster Aurél', 'Gaál Endre', 'Gaál Márta', ' Baróti Tiborné', 'Galamb Sándor', 'Gausz András', 'Gazdapusztai Gyula', 'Gerő Ernőné 1. Fazekas Erzsébet', 'Göncz Lajos', 'Grezsa Ferenc', 'GundaBéla', 'Gyenge Zoltán', 'Gyimesi Sándor', 'Hahn István', 'Hajdú Péter', 'Hajnóczi Gábor', 'Halasy-Nagy József', 'Halász Előd', 'Halász Élődné 1. Szász Anna Mária', 'Hegyi András', 'Helembai Kornélia', 'Heller Ágnes', 'Hermann István Egyed', 'Herrmann Antal', 'Hoffmann Zsuzsanna', 'Horger Antal', 'Hornyánszky Gyula', 'Horváth István Károly', 'Horváth Károly', 'Huszti Dénes', 'Huszti József', 'Ilia Mihály', 'Imre Sándor', 'Imrényi Tibor', 'Ivanics Mária', 'Jakócs Dániel', 'Juhász Antal', 'Juhász Jenő József', 'Juhász József', 'Kanyó Zoltán', 'Kaposi Márton', 'Kardos (Pánd', 'Kari K.Lajos', 'Károly Sándor', 'Karsai László', 'Kecskeméti Ármin', 'Kékes Szabó Mihály', 'Kelemen János', 'Kenesei István', 'Kerényi Károly', 'Keserű Bálint', 'Király István', 'Kiss Lajos', 'Kissné 1. Nóvák Éva', 'Klemm Imre Antal', 'Kocziszky Éva Siflisné', 'Kocsis Mihály', 'Kocsondi András', 'Koltay-Kastner Jenő', 'Komlósy Ákos', 'Koncz János', 'Kontra Miklos', 'Kovács Ilona', 'Kozáky István', 'Kretzoi Sarolta', 'Kristó Gyula', 'Kukovecz Györgyné 1. Zentai Mária', 'Kürtösi Katalin', 'Lagzi Istvan', 'Lepahin Valerij', 'Madácsy László', 'Magyari Zoltánné 1. Techert Margit', 'Makk Ferenc', 'Mályusz Elemér', 'Marinovich Sarolta Resch Béláné', 'Marjanucz Lászlo', 'Márki Sándor', 'Maróti Egon', 'Martonyi Éva', 'Márvány János', 'Masát András', 'Meleczky Márta Judit', 'Mérei Gyula', 'Mester János', 'Mészáros Edit', 'Mészöly Gedeon', 'Mészöly Gedeon', 'Mezősi Károly', 'Mikola Tibor', 'Miskolczy István', 'Módi Mihály', 'Mokány Sándor', 'Nacsády József', 'Nagy Géza', 'Nagy József', 'Nagy László', 'Nagy László J.', 'Nagy Mária', ' Nagy Miklósné', 'Németh T. Enik', 'Nóvák Éva', ' Kissné', 'Nyíri Antal', 'Odorics Ferenc', 'Olajos Terézia', 'Olasz Sándor', 'Orosz Sándor', 'Oroszlán Zoltán', 'Ortutay Gyula', 'Ördögh Éva', 'Ötvös Péter', 'Pál József', 'Pálfy Miklós', 'Pándi Lajos', 'Párducz Mihály', 'Pável Ágoston', 'Penke Olga', ' Penke Botondné', 'Pete István', 'Péter László', 'Pordány László', 'Pósa Péter', 'Pukánszky Béla', 'Rácz Endre', 'Raffay Ernő', 'Resch Béláné 1.Marinovich Sarolta', 'Róna-Tas András', 'Roska Márton', 'Rózsa Éva', 'Rozsnyai Bálint', 'Rubinyi Mózes', 'Sajti Enikő', ' A.', 'Salyámosy Miklós', 'Siflisné 1. Kocziszky Éva', 'Szabó József (Magyar Nyelvészeti T', 'Szabó Tibor', 'Szádeczky-Kardoss Lajos', 'Szádeczky-Kardoss Samu', 'Szajbély Mihály', 'Szakáll Zsigmond', 'Szalamin Edit', 'Szalma Józsefné', ' Vihter Natalia', 'Szántó Imre', 'Szász Anna Mária', ' Halász Élődné', 'Szathmári István', 'Szauder József', 'Székely Lajos', 'Széles Klára', 'Szentiványi Róbert', 'Szerb Antal', 'Szigeti Lajos Sándor', 'Tar Ibolya', 'Techert Margit', ' Magyari Zoltánné', 'Tettamanti Béla', 'Timár Kálmán', 'Tóth Dezső', 'Tóth Imre H.', 'Tóth János', 'Tóth Sándor László', 'Trencsényi-Waldapfel Imre', 'Trogmayer Ottó', 'Varga Ilona', 'Veczkó József', 'Velcsov Mártonné 1. Tóth Katalin', 'Vértes O. József', 'Vidákovich Tibor', 'Vihter Natalia 1. Szalma Józsefné', 'Visy József', 'Vörös László', 'Zentai Mária', ' Kukovecz Györgyné', 'Zimonyi István', 'Zolnai Béla', 'Imre Sándor', 'Ladányi Gedeon', 'Szamosi  János', 'Szász Béla', 'Finaly Henrik', 'Hómao Ottó', 'Hómao Ottó', 'Szász Béla', 'Terner Adolf', 'Szabó Károly', 'Felméri Lajos', 'Szilasi Gergely', 'Hómao Ottó', 'Szamosi  János', 'Szász Béla', 'Hegedűs István', 'Meltzl Hugó', 'Hegedűs István', 'Szinnyei József', 'Schilling Lajos', 'Moldován Gergely', 'Széchy Károly', 'Pecz Vilmos', 'Szádeczky Lajos', 'Márki Sándor', 'Halász Ignác', 'Schneller István', 'Haraszti Gyula', 'Csngery János', 'Böhm Károly', 'Moldován Gergely', 'Vajda Gyula', 'Posta Béla', 'Schilling Lajos', 'Szádeczky Lajos', 'Márki Sándor', 'Schneller István', 'Haraszti Gyula', 'Posta Béla', 'Cholnoky Jenő', 'Zolnai Gyula', 'Dézsi Lajos', 'Schmidt Henrik', 'Zolnai Gyula', 'Erdélyi László', 'Hornyánszky Gyula', 'Finály Henrik Lajos']
Ugyanaz = 0

rows = 0
#Sorok felszámlálása
for row in sheet["D2":"D100000"]:
    for cell in row:
        rows += 1

#Név oszlop törlése
for i in range(2,rows+2):
    if (sheet[f"G{i}"].value != None):
        sheet[f"G{i}"].value = None
starttime = time.time()

#Szó számláló funkció
def szoszamlalo(stri):
    szavak = []
    szo = ""
    for c in stri:
        if c != " ":
            szo += c
        else:
            szavak.append(szo)
            szo = ""
    szavak.append(szo)
    #print(szavak)
    return len(szavak)
#ORIGINAL NLP LOOP
# for i in range(2,rows+2):
#     if (sheet[f"D{i}"].value != None):
#         d = nlp(sheet[f"D{i}"].value)
#         for h in d.ents:
#             if h.label_ == "PER":
#                 sheet[f"G{i}"].value = h.text.title()
#                 break
# print(time.time() - starttime)
#
#NLP TEST FOR 1 CELL 2023.02.21. 13:45 (HIBÁS)
# d = nlp(sheet[f"D20"].value)
# for h in d.ents:
#     print (h, h.label_)
#     if h.label_ == "PER":
#         if (((h.text).find(" ") and len(h.text) > 4)):
#             if ((h.text).find("ny.") < 0):
#                 not_num = True
#                 for c in h.text:
#                     if not c.isalpha():
#                         not_num = False
#                         break
#                 if not_num:
#                     print(h.text)
#                 else:
#                     break

#NLP TEST FOR 1 CELL
# d = nlp(sheet[f"D17"].value)
# for h in d.ents:
#     print (h, h.label_)
#     if h.label_ == "PER":
#         if ((szoszamlalo(h.text) > 1 and len(h.text) > 4)):
#             if ((h.text).find("ny.") < 0):
#                 not_num = True
#                 for c in h.text:
#                     if not c.isalpha():
#                         not_num = False
#                         break
#                 if not_num:
#                     print(h.text)
#                 else:
#                     break

# for i in range(2,rows+2):
#     if (sheet[f"D{i}"].value != None):
#         d = nlp(sheet[f"D{i}"].value)
#         for h in d.ents:
#             print (h, h.label_)
#             if h.label_ == "PER":
#                 if (((h.text).find(" ") and len(h.text) > 4)):
#                     if ((h.text).find("ny.") < 0):
#                         not_num = True
#                         for c in h.text:
#                             if not c.isalpha():
#                                 not_num = False
#                                 break
#                         if not_num:
#                             print(h.text.title())
#                             sheet[f"G{i}"].value = h.text.title()
#                         else:
#                             break

#Eddig jó 2023.02.21. 14:23
# elozotanar = ""
# for i in range(2,rows+2):
#     if (sheet[f"D{i}"].value != None):
#         d = nlp(sheet[f"D{i}"].value)
#         for h in d.ents:
#             #print (h, h.label_)
#             print(str.lower((sheet[f"D{i}"].value)).rfind("ugyan"))
#             if str.lower((sheet[f"D{i}"].value)).rfind("ugyan") >= 0:
#                 print(i)
#                 sheet[f"G{i}"].value = elozotanar
#                 break
#             if h.label_ == "PER":
#                 if ((szoszamlalo(h.text) > 1 and len(h.text) > 4)):
#                     if ((h.text).find("ny.") < 0):
#                         not_num = True
#                         for c in h.text:
#                             if c.isdigit():
#                                 not_num = False
#                                 break
#                         if not_num:
#                             sheet[f"G{i}"].value = h.text.title()
#                             elozotanar = h.text.title()
#                         else:
#                             break


#MAIN loop
elozotanar = ""
for i in range(2,rows+2):
    if i %100 == 0:
        print(i)
    if (sheet[f"D{i}"].value != None):
        d = nlp(sheet[f"D{i}"].value)
        for h in d.ents:
            if h.label_ == "PER":
                if ((szoszamlalo(h.text) > 1 and len(h.text) > 9)):
                    if ((h.text).find("ny.") < 0):
                        not_num = True
                        for c in h.text:
                            if c.isdigit():
                                not_num = False
                                break
                        if not_num:
                            sheet[f"G{i}"].value = h.text.title()
                            elozotanar = sheet[f"G{i}"].value
        if sheet[f"G{i}"].value != None and str.lower(sheet[f"G{i}"].value).rfind("^ny")>-1:
            sheet[f"G{i}"].value = sheet[f"G{i}"].value[0:str.lower(sheet[f"G{i}"].value).rfind("^ny")]
            elozotanar = sheet[f"G{i}"].value
        if (sheet[f"G{i}"].value == "" or sheet[f"G{i}"].value == None):
            if str.lower(sheet[f"D{i}"].value).rfind("ugyanaz") > -1:
                sheet[f"G{i}"].value = elozotanar
                Ugyanaz +=1
                pass
            found = False
            for pos in tanarnevek:
                if str.lower(sheet[f"D{i}"].value).rfind(str.lower(pos)) > -1:
                    sheet[f"G{i}"].value = pos
                    found = True
                    break
            if not found:
                sheet[f"G{i}"].value = elozotanar
                Ugyanaz +=1

print(time.time() - starttime)
print(f"Ugyanazon tanárok: {Ugyanaz}")
book.save(DATA_PATH)