import os
from dadata import Dadata
import polars as pl
# import numpy as np
# from thefuzz import fuzz
from tqdm import tqdm
from dotenv import load_dotenv
load_dotenv()


DADATA_KEY = os.environ["DADATA_KEY"]
SIMILARITY_THRESHOLD = 90

def normalize_list_length(lst, target_length):
    """Расширяет список до заданной длины пустыми значениями."""
    return lst + [""] * (target_length - len(lst))

def merge_excel(df1_name: str, df2_name: str) -> pl.DataFrame:
    df1 = pl.read_excel(df1_name, infer_schema_length=0)
    df2 = pl.read_excel(df2_name, infer_schema_length=0) if df2_name else pl.DataFrame()
    new_df = pl.concat([df1, df2], how='horizontal')
    return new_df
def dadata_test():
    with Dadata(DADATA_KEY) as dadata:
        suggestions = dadata.suggest("party", "054702968779")
        print(suggestions)
        print(suggestions.__len__() > 0)
if __name__ == "__main__":
    dadata_test()
def check_by_inn(df: pl.DataFrame) -> pl.DataFrame:
    print("Начало функции check_by_inn")
    names = []
    for row in tqdm(df.rows(named=True)):
        inn: str = row['ИНН']
        ogrn: str = row.get('ОГРН', '')
        with Dadata(DADATA_KEY) as dadata:
            suggestions = dadata.suggest("party", inn + " " + ogrn)
            match inn.__len__():
                case 12:
                    value: str = suggestions[0]['value'] if suggestions.__len__() > 0 else "ИП "
                    names.append(value.removeprefix('ИП '))
                case 10:
                    value: str = suggestions[0]['data']['management']['name']
                    print(value)
                    print(suggestions[0]['data']['management'])
                    names.append(value)
                case _:
                    names.append("")
    df = df.with_columns(pl.Series(name='ФИО', values=names))
    return df

