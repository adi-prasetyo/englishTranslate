import pandas as pd
import re
import os
import unicodedata

folderPath = r"D:\PycharmProjects\englishTranslate\ingredients"

ingredients = "ingredients.xlsx"
ingredients_dictionary = "ingredients_dictionary.xlsx"
ingredients_separator = "ingredients_separator.xlsx"
result = "result.xlsx"
ingredients_unknown = "ingredients_unknown.xlsx"

df_ingredients = pd.read_excel(os.path.join(folderPath, ingredients))
df_dictionary = pd.read_excel(os.path.join(folderPath, ingredients_dictionary))
df_separator = pd.read_excel(os.path.join(folderPath, ingredients_separator))

# put into dict
sepDict = dict(zip(df_separator.Separator_original, df_separator.Separator_clean))
translationDict = dict(zip(df_dictionary.Japanese, df_dictionary.English))

# separate original and japanese modified for easier debugging
df_ingredients["ingredients"] = df_ingredients["ingredients_ori"].str.strip()

# normalize the numbers etc, remove all spaces and clean the separator
df_ingredients["ingredients"] = df_ingredients["ingredients"].str.normalize("NFKC")
df_ingredients["ingredients"] = [
    "".join(re.split(" ", e)) for e in df_ingredients["ingredients"]
]
df_ingredients["ingredients"] = [
    "".join(sepDict.get(item, item) for item in re.split("(\W)", e))
    for e in df_ingredients["ingredients"]
]

# replace japanese with english dict
df_ingredients["result"] = [
    "".join(translationDict.get(item, item) for item in re.split("(\W)", e))
    for e in df_ingredients["ingredients"]
]
# replaced twice cuz stupid '-' separator so need to replace it again with comma split
df_ingredients["result"] = [
    ", ".join(translationDict.get(item, item) for item in re.split(", ", e))
    for e in df_ingredients["result"]
]

# add space before brackets and after commas
df_ingredients["result"] = [
    re.sub(r"(\S)\(", r"\1 (", e) for e in df_ingredients["result"]
]
df_ingredients["result"] = [
    re.sub(r"(?<=[.,:])(?=[^\s])", r" ", e) for e in df_ingredients["result"]
]

# capitalize first letter
df_ingredients["result"] = [
    myString[:1].upper() + myString[1:] for myString in df_ingredients["result"]
]

df_ingredients.to_excel(
    os.path.join(folderPath, result), sheet_name="Sheet1", index=False
)

# get the japanese unknown words
flat_list = [
    item
    for sublist in [re.split("\W+", e) for e in df_ingredients["result"]]
    for item in sublist
]

jap_word = set()

for word in flat_list:
    if word.isascii() == False:
        jap_word.add(word)
df_jap = pd.DataFrame(list(jap_word))
df_jap.to_excel(
    os.path.join(folderPath, ingredients_unknown), sheet_name="Sheet1", index=False
)
# finished
