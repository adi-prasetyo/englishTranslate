import pandas as pd
import re
import unicodedata
from googletrans import Translator
import os

def rreplace(s, occurrence, old, new):
    li = s.rsplit(old, occurrence)
    return new.join(li)

def replace_endswith(text, **kwargs):
    result = False
    for key, value in kwargs.items():
        if text.endswith(key):
            result = rreplace(text, 1, key, value)        
            return result
    if result == False:
        return text
    
def replace_startswith(text, **kwargs):
    result = False
    for key, value in kwargs.items():
        if text.startswith(key):
            result = text.replace(key, value, 1)        
            return result
    if result == False:
        return text 

folderPath = r'C:\Users\adipr\englishNameTranslate'

dictionary_space = 'dictionary_space.xlsx'
dictionary_endswith = 'dictionary_endswith.xlsx'
dictionary_startswith = 'dictionary_startswith.xlsx'
dictionary_replace = 'dictionary_replace.xlsx'

df_space = pd.read_excel(os.path.join(folderPath, dictionary_space))
df_endswith = pd.read_excel(os.path.join(folderPath, dictionary_endswith))
df_startswith = pd.read_excel(os.path.join(folderPath, dictionary_startswith))
df_replace = pd.read_excel(os.path.join(folderPath, dictionary_replace))

df_space.sort_values(by=['English'], ascending=False, inplace=True)
df_startswith.sort_values(by=['English'], ascending=False, inplace=True)
df_replace.sort_values(by=['English'], ascending=False, inplace=True)

df_replace['English'] = df_replace['English'] + " " #add space to the end cuz there's no space in the jap version
df_startswith['English'] = df_startswith['English'] + " " #add space to the end cuz there's no space in the jap version
df_space['English'] = " " + df_space['English'] + " " #add space to the end cuz there's no space in the jap version

templatePath = r'C:\Users\adipr\Documents\Excel\template'
japanese_file = 'translate eng.xlsx'
df_japanese = pd.read_excel(os.path.join(templatePath, japanese_file))

spaceDict = dict(zip(df_space['Japanese'], df_space['English']))
replaceDict = dict(zip(df_replace['Japanese'], df_replace['English']))
startswithDict = dict(zip(df_startswith['Japanese'], df_startswith['English']))
endswithDict = dict(zip(df_endswith['Japanese'], df_endswith['English']))

df_result = df_japanese.copy()

jap_col = 'Product Name'
eng_col = 'English'

df_result[jap_col] = df_result[jap_col].str.replace('＄', '', regex=True)
df_result[jap_col] = df_result[jap_col].str.normalize('NFKC')
df_result[jap_col] = df_result[jap_col].str.strip() #strip leading and trailing white spaces
df_result[eng_col] = df_result[eng_col].astype(str)

df_result[eng_col] = df_result.apply(lambda x: replace_startswith(x[jap_col], **startswithDict), axis=1)
df_result[eng_col] = df_result.apply(lambda x: replace_endswith(x[eng_col], **endswithDict), axis=1)
df_result[eng_col] = [' '.join(spaceDict.get(item,item) for item in re.split(' ', e)) for e in df_result[eng_col]]

df_result[eng_col] = df_result[eng_col].replace(replaceDict, regex=True)

df_result[eng_col] = df_result[eng_col].str.replace('\s+', ' ', regex=True) #remove multiple spaces
df_result[eng_col] = df_result[eng_col].str.strip() #remove multiple spaces

df_result['Barcode Number'] = df_result['Barcode Number'].astype(str)

result = 'result.xlsx'
writer = pd.ExcelWriter(os.path.join(folderPath, result), engine='xlsxwriter')
df_result.to_excel(writer,  sheet_name='Sheet1', index = False)  

worksheet = writer.sheets['Sheet1']

# Set the column width and format.
worksheet.set_column('A:A', 6)
worksheet.set_column('B:K', 0)
worksheet.set_column('N:O', 0)
worksheet.set_column('S:S', 0)
worksheet.set_column('M:M', 0)
worksheet.set_column('L:L', 15) #JAN column
worksheet.set_column('P:P', 25) #maker column
worksheet.set_column('Q:Q', 60) #japanese column
worksheet.set_column('R:R', 90) #english column
worksheet.set_column('T:AS', 0) #leave the last column shown so can copy paste easily

writer.save()