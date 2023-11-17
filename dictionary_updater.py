# %%
import pandas as pd

from datetime import datetime
from df_config import create_postgres_engine, create_aws_engine

aws_engine = create_aws_engine()
postgres_engine = create_postgres_engine()

df_ingredients_dictionary = pd.read_excel(r'D:\PycharmProjects\englishTranslate\ingredients\ingredients_dictionary.xlsx')
df_ingredients_separator = pd.read_excel(r'D:\PycharmProjects\englishTranslate\ingredients\ingredients_separator.xlsx')

df_ingredients_dictionary.to_sql('ingredients_dictionary', postgres_engine, if_exists='replace', index=False)
df_ingredients_dictionary.to_sql('ingredients_dictionary', aws_engine, if_exists='replace', index=False)

df_ingredients_separator.to_sql('ingredients_separator', postgres_engine, if_exists='replace', index=False)
df_ingredients_separator.to_sql('ingredients_separator', aws_engine, if_exists='replace', index=False)

# Update the dictionary_update_log table
update_time = datetime.utcnow()
df_log_update = pd.DataFrame({
    'dictionary_name': ['ingredients_dictionary', 'ingredients_separator'],
    'last_updated': [update_time, update_time]
})
df_log_update.to_sql('dictionary_update_log', postgres_engine, if_exists='replace', index=False)
df_log_update.to_sql('dictionary_update_log', aws_engine, if_exists='replace', index=False)


