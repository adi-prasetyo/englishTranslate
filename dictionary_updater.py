# %%
import pandas as pd

from datetime import datetime
from sqlalchemy import create_engine

def create_postgres_engine():
    return create_engine("postgresql+psycopg2://postgres:lalala@localhost:5432/postgres")


def create_aws_engine():
    return create_engine('postgresql+psycopg2://postgres:taji3030postgres@tajimaya-products-instance-1.crjzkl0txyjn.ap-northeast-1.rds.amazonaws.com:5432/postgres')

aws_engine = create_aws_engine()
postgres_engine = create_postgres_engine()

df_ingredients_dictionary = pd.read_excel(r'D:\PycharmProjects\englishTranslate\ingredients\ingredients_dictionary.xlsx')
df_ingredients_separator = pd.read_excel(r'D:\PycharmProjects\englishTranslate\ingredients\ingredients_separator.xlsx')

# Update the dictionary_update_log table
update_time = datetime.utcnow()
df_log_update = pd.DataFrame({
    'dictionary_name': ['ingredients_dictionary', 'ingredients_separator'],
    'last_updated': [update_time, update_time]
})

engines = [aws_engine, postgres_engine]

for engine in engines:
    df_ingredients_dictionary.to_sql('ingredients_dictionary', engine, if_exists='replace', index=False)
    df_ingredients_separator.to_sql('ingredients_separator', engine, if_exists='replace', index=False)
    df_log_update.to_sql('dictionary_update_log', engine, if_exists='replace', index=False)

