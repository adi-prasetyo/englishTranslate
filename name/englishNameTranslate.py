import pandas as pd
import time
import re
import os
import sys
import shutil
import unicodedata as ud

from titlecase import titlecase
from random import randint
from googletrans import Translator
from datetime import datetime
from dateutil.relativedelta import relativedelta

templatePath = r'D:\Excel\template'
folderPath = r"D:\PycharmProjects\englishTranslate\name"

suffixList = 'Item List English.xlsx'
suffixTranslate = 'Translate Eng.xlsx'

dictionary_space = "dictionary_space.xlsx"
dictionary_endswith = "dictionary_endswith.xlsx"
dictionary_startswith = "dictionary_startswith.xlsx"
dictionary_replace = "dictionary_replace.xlsx"
dictionary_flavor = "dictionary_flavor.xlsx"
dictionary_last = "dictionary_last.xlsx"

name_col = "Product Name"
jap_col = "Japanese"
eng_col = "English"
google_col = "Google Translation"

translator = Translator()

df_endswith = pd.read_excel(os.path.join(folderPath, dictionary_endswith))
df_startswith = pd.read_excel(os.path.join(folderPath, dictionary_startswith))
df_replace = pd.read_excel(os.path.join(folderPath, dictionary_replace))
df_flavor = pd.read_excel(os.path.join(folderPath, dictionary_flavor))
df_last = pd.read_excel(os.path.join(folderPath, dictionary_last))

latin_letters= {}

# normalize the zenkaku and remove all spaces in the jap_col column
def normalize_df(df, col):
        df[col] = df[col].str.normalize("NFKC")
        df[col] = ["".join(re.split(" ", e)) for e in df[col]]


def sort_df(df, col):
    # do this before sorting the length
    normalize_df(df, col)
    df["length"] = df[col].str.len()
    df.sort_values(by=["length", "English"], ascending=False, inplace=True)
    return df


# df dictionary processing
for df in [df_endswith, df_startswith, df_replace, df_flavor, df_last]:
    # sort df and add space
    df = sort_df(df, jap_col)
    df["English"] = " " + df["English"] + " "


replaceDict = dict(zip(df_replace[jap_col], df_replace["English"]))
startswithDict = dict(zip(df_startswith[jap_col], df_startswith["English"]))
endswithDict = dict(zip(df_endswith[jap_col], df_endswith["English"]))
flavorDict = dict(zip(df_flavor[jap_col], df_flavor["English"]))
lastDict = dict(zip(df_last[jap_col], df_last["English"]))

def is_latin(uchr):
    try: return latin_letters[uchr]
    except KeyError:
         return latin_letters.setdefault(uchr, 'LATIN' in ud.name(uchr))


# from stackoverflow, no idea how it works
def only_roman_chars(unistr):
    return all(is_latin(uchr)
           for uchr in unistr
           if uchr.isalpha())


def getFileName(_time):
    year_month_str = _time.strftime("%Y %B")     
    listname = " ".join([year_month_str, suffixList])
    writename = " ".join([year_month_str, suffixTranslate])
    
    outputfile = os.path.join(templatePath, listname)
    writefile = os.path.join(templatePath, writename)
    
    return outputfile, writefile, writename


def concat_all(filetime, eng_col=eng_col):
    
    outputfile, writefile, writename = getFileName(filetime)
    
    df_snacks = pd.read_excel(outputfile, 0)

    # categories are currently not working, so disable these parts
    # df_drinks = pd.read_excel(outputfile, 1)
    # df_foods = pd.read_excel(outputfile, 2)
    
    # df_concat = pd.concat([df_snacks, df_drinks, df_foods])
    # temp solution for now
    df_concat = df_snacks.copy()

    df_concat.rename(columns = {'English Ingredients':google_col}, inplace = True)

    # only select column that has no translation
    # for some reason that column value is 0
    df_concat = df_concat.loc[df_concat[eng_col] == 0]
    
    df_concat.reset_index(drop=True, inplace=True)
    
    return df_concat, writefile, writename


def excel_write(df, writeFile):
    with pd.ExcelWriter(writeFile, engine="xlsxwriter") as writer:        
        df.to_excel(writer, index=False)
        worksheet = writer.sheets["Sheet1"]

        # Set the column width and format.
        worksheet.set_column("A:A", 6)
        worksheet.set_column("B:K", 0)
        worksheet.set_column("L:L", 15)  # JAN column
        worksheet.set_column("M:O", 0)
        worksheet.set_column("P:P", 25)  # maker column
        worksheet.set_column("Q:Q", 60)  # jap_col column
        worksheet.set_column("R:R", 90)  # english column
        worksheet.set_column("S:S", 90)  # google column
        worksheet.set_column("T:AS", 0)
        
    print(writeFile)


# make exception for all the words that are already capitalized
# but only for words here
def abbreviations(word, **kwargs):
    if word.upper() in ('AGF', 'QP', 'SSK', 'YBC', 'QTTA', 'QBB', 'UFO', 'U.F.O'):
        return word.upper() 
    elif word == word.upper():
        return word.capitalize()


def google_translate_col(df, 
                    name_col=name_col, 
                    google_col=google_col,
                    limit=10):

    translation_text = []
    
    for x in df[name_col]:

        # if for some reason the value is 0 then dont translate
        if len(x) < 2:
            translation_text.append("")
            continue
        
        # if the text contains no jap_col chr then return as it is
        if only_roman_chars(x):
            translation_text.append(x)
            continue

        initial_wait = randint(1,3)
        retry_wait = randint(2,7)
        attempts = 1
        success = False

        time.sleep(initial_wait)

        while attempts < limit and not success:
            try:
                trans = translator.translate(x)
                transTitle = titlecase(trans.text, callback=abbreviations)

                translation_text.append(transTitle)        
                print(transTitle)
                success = True
            except:
                attempts += 1
                time.sleep(retry_wait)
                if attempts == limit:
                    print(x + " translation failed")
                    translation_text.append(x)

    # manual correction for Google Translate
    # translation_text = [sub.replace('Pet', 'PET') for sub in translation_text] #Petit becomes PETit so disable this for now
    # add space at the end to make sure these are not part of words
    translation_text = [sub.replace('Ucc ', 'UCC ') for sub in translation_text]
    translation_text = [sub.replace('U.f.o.', 'U.F.O.') for sub in translation_text]
    translation_text = [sub.replace('Agf ', 'AGF ') for sub in translation_text]
    translation_text = [sub.replace('S & B ', 'S&B ') for sub in translation_text]
    translation_text = [sub.replace('AceCook ', 'Acecook ') for sub in translation_text]
    df[google_col] = translation_text


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


def translate_df(df, 
                name_col=name_col, 
                eng_col=eng_col,
                startswithDict=startswithDict, 
                endswithDict=endswithDict,
                replaceDict=replaceDict,
                flavorDict=flavorDict,
                lastDict=lastDict):

    # delete all the shuubai mark
    df[name_col] = df[name_col].str.replace("＄", "", regex=True)

    # copy the original df
    df_result = df.copy()

    normalize_df(df_result, name_col)

    df_result[eng_col] = df_result[eng_col].astype(str)

    df_result[eng_col] = df_result.apply(
        lambda x: replace_startswith(x[name_col], **startswithDict), axis=1
    )
    df_result[eng_col] = df_result.apply(
        lambda x: replace_endswith(x[eng_col], **endswithDict), axis=1
    )

    # replace words regardless everything
    for regexDict in [replaceDict, flavorDict, lastDict]:
        df_result[eng_col] = df_result[eng_col].replace(regexDict, regex=True)
    df_result[eng_col] = df_result[eng_col].str.replace(
        "\s+", " ", regex=True
    )  # remove multiple spaces
    df_result[eng_col] = df_result[
        eng_col
    ].str.strip()  # strip leading and trailing white spaces

    df_result["Barcode Number"] = df_result["Barcode Number"].astype(str)

    # revert to the original product name for checking later
    df_result[name_col] = df[name_col]

    return df_result

# df will be empty if there is nothing to translate
df_thismonth, thisWriteFile, thisFileName = concat_all(datetime.now())
df_nextmonth, nextWriteFile, nextFileName = concat_all(datetime.now() + relativedelta(months=1))

# early exit if col eng are all filled
if len(df_thismonth) == 0 and len(df_nextmonth) == 0:
    sys.exit("All product names are already translated.")

# translate the jap name with custom dict if there is any translation
# translate with google translation
# from the half-translated col, not the original jap col
if len(df_thismonth) != 0:
    df_thismonth_translated = translate_df(df_thismonth)
    google_translate_col(df_thismonth_translated, name_col=eng_col)
    excel_write(df_thismonth_translated, thisWriteFile)

if len(df_nextmonth) != 0:
    df_nextmonth_translated = translate_df(df_nextmonth)
    google_translate_col(df_nextmonth_translated, name_col=eng_col)
    excel_write(df_nextmonth_translated, nextWriteFile)

dropboxDir = r"D:\Dropbox\Excel\Translate"

shutil.copyfile(
            thisWriteFile,
            os.path.join(dropboxDir, thisFileName),
        )

shutil.copyfile(
            nextWriteFile,
            os.path.join(dropboxDir, nextFileName),
        )