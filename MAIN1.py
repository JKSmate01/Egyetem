import spacy
import huspacy
import nltk
import math

szoveg = "3.  Psychologia; (bevezetésül a neveléstanbaj; hétfőn, kedden, szerdán, d. e. 10—1 I óráig. Felméri Lajos ny. r. tanár."

nlp = spacy.load("hu_core_news_lg")

a = nlp(szoveg)
for i in a:
    print(i,i.ent_type_) #probaljuk kiszedni a datet