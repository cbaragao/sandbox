import pandas as pd
import os
import nltk
from nltk.corpus import stopwords
from nltk.tag import pos_tag

df = pd.read_excel (r'C:\...', sheet_name='Sheet1')

df.head()


df['Synopsis'] = df['Synopsis'].str.replace('[^A-Za-z0-9]+', ' ', regex = True)

# Import stopwords with nltk.
stop = stopwords.words('english') 

df['Synopsis'] = df['Synopsis'].str.lower()

df['Synopsis'] = df['Synopsis'].fillna("NULL")

df['Synopsis'] = df['Synopsis'].str.replace('\d+', '', regex=True)

df['Synopsis_clean'] = df['Synopsis'].apply(lambda x: ' '.join([word for word in x.split() if word not in (stop)]))

# for natural language processing: named entity recognition
import spacy
from collections import Counter

import en_core_web_sm
nlp = en_core_web_sm.load()
nlp.max_length = 2000000
 
tokens = nlp(''.join(str(df.Synopsis_clean.tolist())))

#-----------------------------------------
#GPE = Countries, cities, states

gpe_list = []

for ent in tokens.ents:
   if ent.label_ == 'GPE':
       gpe_list.append(ent.text)
        
gpe_counts = Counter(gpe_list).most_common(20)
df_gpe = pd.DataFrame(gpe_counts, columns =['text', 'count'])
df_gpe.head(10)

#-----------------------------------------
#LOC = Non-GPE locations, mountain ranges, bodies of water

#loc_list = []

#for ent in tokens.ents:
 #  if ent.label_ == 'LOC':
 #      loc_list.append(ent.text)
        
#loc_counts = Counter(loc_list).most_common(20)
#df_loc = pd.DataFrame(loc_counts, columns =['text', 'count'])
#df_loc.head()

#-----------------------------------------
#ORG = Companies, agencies, institutions, etc

#org_list = []

#for ent in tokens.ents:
#   if ent.label_ == 'ORG':
#       org_list.append(ent.text)
        
#org_counts = Counter(org_list).most_common(20)
#df_org = pd.DataFrame(org_counts, columns =['text', 'count'])
#df_org.head()
#-----------------------------------------

#import nltk
#nltk.download('wordnet')
#from ntlk.corpus import wordnet
#from nltk.stem import WordNetLemmatizer

#w_tokenizer = nltk.tokenize.WhitespaceTokenizer()
#lemmatizer = WordNetLemmatizer()

#def lemmatize_text(text):
    #return [lemmatizer.lemmatize(w) for w in w_tokenizer.tokenize(text)]


#df['Synopsis_lem'] = df.Synopsis.apply(lemmatize_text)

#df.to_csv(r'/Users/...')
