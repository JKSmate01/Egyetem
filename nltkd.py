import nltk
from nltk.tokenize import word_tokenize
from nltk.tag import pos_tag
from openpyxl import load_workbook
nevek = ['Ágoston György', 'Albert Sándor', 'Almási Tibor', 'Anderle Ádám', 'Bakos Ferenc', 'Balázs Mihály', 'Bálint Alajos', 'Bálint Sándor', 'Balogh Tibor', 'Banner János', 'Bánréti Zoltán', 'Baranyai Erzsébet', 'Baranyai Zoltán', 'Bárczi Géza', 'Barna Gábor', 'Baróti Dezső', 'Baróti Tiborné 1. Gaál Márta', 'Bartók György', 'Bassola Péter', 'Belényi Gyula', 'Bellon Tibor', 'Benedek Nándor', 'Bényiné 1. Farkas Mária', 'Bérezik Árpád', 'Bernáth Árpád', 'Berta Árpád', 'Birkás Géza', 'Boda István', 'Bognár Cecil Pál', 'Bollobás Enikő', 'Buday Árpád', 'Czachesz Erzsébet Cs.', 'Csapó Benő', 'Csefkó Gyula', 'Csejtei Dezső','Csengery János', 'Cseresnyési László', 'Csetri Lajos', 'Csukás István', 'Csúri Károly', 'Deér József', 'Deme László', 'Dézsi Lajos', 'Domokos Péter', 'Dornbach Mária', 'Duró Lajos', 'Eördögh István', 'Eperjessy Kálmán', 'Erdélyi Gyula', 'Erdélyi László', 'Erdélyi László Gyula', 'Erős Ferenc', 'Fabiny Tibor', 'Fábricz Károly', 'Farkas Mária', ' Bényiné', 'Fazekas Erzsébet', ' Gerő Ernőné', 'Fejér Ádám', 'Fejes Katalin', ' B.', 'Feleky Gábor', 'Felvinczi Takáts Zoltán', 'Fenyvesi István', 'Ferenczi Imre', 'Fodor István', 'Fogarasi Miklós', 'Fógel József', 'Fónagy Iván', 'Forgács Tamás', 'Förster Aurél', 'Gaál Endre', 'Gaál Márta', ' Baróti Tiborné', 'Galamb Sándor', 'Gausz András', 'Gazdapusztai Gyula', 'Gerő Ernőné 1. Fazekas Erzsébet', 'Göncz Lajos', 'Grezsa Ferenc', 'GundaBéla', 'Gyenge Zoltán', 'Gyimesi Sándor', 'Hahn István', 'Hajdú Péter', 'Hajnóczi Gábor', 'Halasy-Nagy József', 'Halász Előd', 'Halász Élődné 1. Szász Anna Mária', 'Hegyi András', 'Helembai Kornélia', 'Heller Ágnes', 'Hermann István Egyed', 'Herrmann Antal', 'Hoffmann Zsuzsanna', 'Horger Antal', 'Hornyánszky Gyula', 'Horváth István Károly', 'Horváth Károly', 'Huszti Dénes', 'Huszti József', 'Ilia Mihály', 'Imre Sándor', 'Imrényi Tibor', 'Ivanics Mária', 'Jakócs Dániel', 'Juhász Antal', 'Juhász Jenő József', 'Juhász József', 'Kanyó Zoltán', 'Kaposi Márton', 'Kardos (Pánd', 'Kari K.Lajos', 'Károly Sándor', 'Karsai László', 'Kecskeméti Ármin', 'Kékes Szabó Mihály', 'Kelemen János', 'Kenesei István', 'Kerényi Károly', 'Keserű Bálint', 'Király István', 'Kiss Lajos', 'Kissné 1. Nóvák Éva', 'Klemm Imre Antal', 'Kocziszky Éva Siflisné', 'Kocsis Mihály', 'Kocsondi András', 'Koltay-Kastner Jenő', 'Komlósy Ákos', 'Koncz János', 'Kontra Miklos', 'Kovács Ilona', 'Kozáky István', 'Kretzoi Sarolta', 'Kristó Gyula', 'Kukovecz Györgyné 1. Zentai Mária', 'Kürtösi Katalin', 'Lagzi Istvan', 'Lepahin Valerij', 'Madácsy László', 'Magyari Zoltánné 1. Techert Margit', 'Makk Ferenc', 'Mályusz Elemér', 'Marinovich Sarolta Resch Béláné', 'Marjanucz Lászlo', 'Márki Sándor', 'Maróti Egon', 'Martonyi Éva', 'Márvány János', 'Masát András', 'Meleczky Márta Judit', 'Mérei Gyula', 'Mester János', 'Mészáros Edit', 'Mészöly Gedeon', 'Mészöly Gedeon', 'Mezősi Károly', 'Mikola Tibor', 'Miskolczy István', 'Módi Mihály', 'Mokány Sándor', 'Nacsády József', 'Nagy Géza', 'Nagy József', 'Nagy László', 'Nagy László J.', 'Nagy Mária', ' Nagy Miklósné', 'Németh T. Enik', 'Nóvák Éva', ' Kissné', 'Nyíri Antal', 'Odorics Ferenc', 'Olajos Terézia', 'Olasz Sándor', 'Orosz Sándor', 'Oroszlán Zoltán', 'Ortutay Gyula', 'Ördögh Éva', 'Ötvös Péter', 'Pál József', 'Pálfy Miklós', 'Pándi Lajos', 'Párducz Mihály', 'Pável Ágoston', 'Penke Olga', ' Penke Botondné', 'Pete István', 'Péter László', 'Pordány László', 'Pósa Péter', 'Pukánszky Béla', 'Rácz Endre', 'Raffay Ernő', 'Resch Béláné 1.Marinovich Sarolta', 'Róna-Tas András', 'Roska Márton', 'Rózsa Éva', 'Rozsnyai Bálint', 'Rubinyi Mózes', 'Sajti Enikő', ' A.', 'Salyámosy Miklós', 'Siflisné 1. Kocziszky Éva', 'Szabó József (Magyar Nyelvészeti T', 'Szabó Tibor', 'Szádeczky-Kardoss Lajos', 'Szádeczky-Kardoss Samu', 'Szajbély Mihály', 'Szakáll Zsigmond', 'Szalamin Edit', 'Szalma Józsefné', ' Vihter Natalia', 'Szántó Imre', 'Szász Anna Mária', ' Halász Élődné', 'Szathmári István', 'Szauder József', 'Székely Lajos', 'Széles Klára', 'Szentiványi Róbert', 'Szerb Antal', 'Szigeti Lajos Sándor', 'Tar Ibolya', 'Techert Margit', ' Magyari Zoltánné', 'Tettamanti Béla', 'Timár Kálmán', 'Tóth Dezső', 'Tóth Imre H.', 'Tóth János', 'Tóth Sándor László', 'Trencsényi-Waldapfel Imre', 'Trogmayer Ottó', 'Varga Ilona', 'Veczkó József', 'Velcsov Mártonné 1. Tóth Katalin', 'Vértes O. József', 'Vidákovich Tibor', 'Vihter Natalia 1. Szalma Józsefné', 'Visy József', 'Vörös László', 'Zentai Mária', ' Kukovecz Györgyné', 'Zimonyi István', 'Zolnai Béla', 'Imre Sándor', 'Ladányi Gedeon', 'Szamosi  János', 'Szász Béla', 'Finaly Henrik', 'Hómao Ottó', 'Hómao Ottó', 'Szász Béla', 'Terner Adolf', 'Szabó Károly', 'Felméri Lajos', 'Szilasi Gergely', 'Hómao Ottó', 'Szamosi  János', 'Szász Béla', 'Hegedűs István', 'Meltzl Hugó', 'Hegedűs István', 'Szinnyei József', 'Schilling Lajos', 'Moldován Gergely', 'Széchy Károly', 'Pecz Vilmos', 'Szádeczky Lajos', 'Márki Sándor', 'Halász Ignác', 'Schneller István', 'Haraszti Gyula', 'Csngery János', 'Böhm Károly', 'Moldován Gergely', 'Vajda Gyula', 'Posta Béla', 'Schilling Lajos', 'Szádeczky Lajos', 'Márki Sándor', 'Schneller István', 'Haraszti Gyula', 'Posta Béla', 'Cholnoky Jenő', 'Zolnai Gyula', 'Dézsi Lajos', 'Schmidt Henrik', 'Zolnai Gyula', 'Erdélyi László', 'Hornyánszky Gyula', 'Finály H. Lajos']
book = load_workbook('ossz.xlsx',data_only=True)
sheet = book["Egybe"]
def funkcijó(stri):
    e = []
    szavak = word_tokenize(stri)
    tagok = pos_tag(szavak)
    for i in tagok:
        if i[1] == "NNP":
            e.append(i[0])
    return szavak

#Működőképes
# def nagybetusszoszetvalaszto(stri):
#     nevszavak = []
#     szo = ""
#     for c in stri:
#         if c != " " and c.isupper():
#             szo += c
#         else:ss
#             nevszavak.append(szo)
#             szo = ""
#     print(nevszavak)
#     return nevszavak


def nagybetusszoszetvalaszto(stri):
    nevszavak = []
    szo = ""
    for c in stri:
        if c != " ":
            szo += c
        else:
            if ((szo.isalpha() and szo.isupper()) and len(szo) >= 3) or szo == "Dr.":
                nevszavak.append(szo)
            szo = ""
    print(nevszavak)
    return nevszavak

#print(funkcijó("Jakus Máté MagyaroRszág jógyerek"))
rows = 0
for row in sheet["D2":"D100000"]:
    for cell in row:
        rows += 1

#Szavak NLTK próba 3 Jakus
# elozonev = None
# megvan = False
# for i in range(2, rows+2):
#     if (sheet[f"D{i}"].value != None):
#         nltktab = funkcijó(sheet[f"D{i}"].value)
#         for d in range(len(nltktab)):
#             for n in range(len(nevek)):
#                 if((nltktab[d] in nevek[n]) and len(nltktab[d]) > 4):
#                     print(nltktab[d],nevek[n])
#                     sheet[f"G{i}"].value = nevek[n]
#                     break
#             break
#END

#Szavak NLTK próba 4 Martin
# for i in range(2, rows+2):
#     if (sheet[f"D{i}"].value != None):
#         nltktab = funkcijó(sheet[f"D{i}"].value)
#         for d in range(len(nltktab)):
#             for n in range(len(nevek)):


                    #print(nevek[n])
#END


#END
#Időpont
for d in range(2,rows+2):
    if (sheet[f"D{d}"].value != None):
        for i in range(len(sheet[f"D{d}"].value)):
            if (sheet[f"D{d}"].value[i] == ";"):
                next = sheet[f"D{d}"].value[i+1:]
                if ("ig." in next):
                    ds = next[:next.rfind("ig.")]
                break
        sheet[f"E{d}"].value = sheet[f"D{d}"].value[:i]
        sheet[f"F{d}"].value = ds + "ig."
sheet["E2"].value = funkcijó(sheet["D2"].value)[0]
#END

for d in range(2,rows+2):
    if (sheet[f"D{d}"].value != None):
        for i in range(len(sheet[f"D{d}"].value)):
            if (sheet[f"D{d}"].value[i] == ";"):
                next = sheet[f"D{d}"].value[i+1:]
                if ("ig." in next):
                    ds = next[:next.rfind("ig.")]
                    sheet[f"G{d}"].value = sheet[f"D{d}"].value[sheet[f"D{d}"].value.rfind("ig.")+3:]
                break
        sheet[f"E{d}"].value = sheet[f"D{d}"].value[:i]
        sheet[f"F{d}"].value = ds + "ig."
sheet["E2"].value = funkcijó(sheet["D2"].value)[0]
#END
for i in range(2, rows+2):
    megvan = False
    if (sheet[f"D{i}"].value != None):
        nltktab = funkcijó(sheet[f"D{i}"].value)
        for t in range(len(nltktab)-1):
            for k in range(len(nevek)-1):
                if nltktab[t-1]+" "+nltktab[t] == nevek[k]:
                    sheet[f"G{i}"].value = nevek[k]
                    elozonev = nevek[k]
                    megvan = True
                    break
            if megvan:
                break
#Nagybetűs nevek
for i in range(2, rows+2):
    if (sheet[f"D{i}"].value != None):
        szavakszetvalasztva = nagybetusszoszetvalaszto(sheet[f"D{i}"].value)
        try:
            if len(szavakszetvalasztva) > 1:
                szavakstring = ""
                for szo in range(len(szavakszetvalasztva)):
                    if szo > 0:
                        szavakstring += " "+szavakszetvalasztva[szo].title()
                    else:
                        szavakstring += szavakszetvalasztva[szo].title()
                
                sheet[f"G{i}"].value = szavakstring
                
                elozonev = szavakstring
        except:
            pass
for n in range(2, rows+2):
    if (sheet[f"G{n}"].value != None):
        for ne in range(len(nevek)):
            if nevek[ne] in sheet[f"G{n}"].value:
                sheet[f"G{n}"].value = nevek[ne]
    
book.save('.\egyetem\ossz.xlsx')

#29.Magyar hatások a román irodalomra. (Folyta¬tás.) Heti 2 óra. Szombaton d. e. 8—10-ig. Dr. SULICA SZILÁRD megbízott előadó, a IV. sz. tanteremben.
