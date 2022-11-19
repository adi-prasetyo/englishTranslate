import os
import pandas as pd
import shutil

folderPath = r'C:\Users\adipr\Documents\Excel'

excelname = input("Please enter Excel name (without extension. And xlsx only):")

excelfile = excelname + ".xlsx"

df = pd.read_excel(os.path.join(folderPath, excelfile), dtype=str)
df['JANjpg'] = df.iloc[:, 11] + '.jpg'
df = df.dropna(subset=['JANjpg'])
df['JANjpg'] = df['JANjpg'].apply('{:0>17}'.format)

save_imagePath = r'D:/save_image/'
japonPath = r'D:/japon/'
emptylist = []

for filename in df['JANjpg']:
    try:
        shutil.copyfile(os.path.join(save_imagePath, filename), os.path.join(japonPath, filename))
    except:
        emptylist.append(filename[:-4]) #remove the '.jpg' from the list so can select the JAN easier on text file

print("Empty List length: " + str(len(emptylist)))

textFile = 'emptylist.txt'

with open(os.path.join(japonPath, textFile), 'w') as f:
    for item in emptylist:
        f.write("%s\n" % item)