import huspacy
import spacy
import hu_core_news_lg
nlphucore = hu_core_news_lg.load()
nlpspacy = spacy.load("hu_core_news_lg")
nlphuspacy = huspacy.load()

nlphuscore = nlphucore("27. A latin értekezés stílusa. (Gyakorlatokkal.) Heti 2 óra.Később megállapítandó időben és helyen. Dr. Marót Károly lector.")
nlphusd = nlphuspacy("27. A latin értekezés stílusa. (Gyakorlatokkal.) Heti 2 óra.Később megállapítandó időben és helyen. Dr. Marót Károly lector.")
nlphucored = nlpspacy("27. A latin értekezés stílusa. (Gyakorlatokkal.) Heti 2 óra.Később megállapítandó időben és helyen. Dr. Marót Károly lector.")
print(nlphusd)
print(nlphucored)
print(nlphuscore)