import pandas as pd
import re

catalog_file = 'ENST Library Catalog.xlsx'
codes_file = 'lc-codes.xlsx'

df = pd.read_excel(catalog_file)
lc_codes_df = pd.read_excel(codes_file)

def get_lc_subclass(lc_code):
    lc_main_class = ''.join(filter(str.isalpha, lc_code))[:2]

    if lc_main_class in ['BF', 'LB', 'LC', 'PQ']:
        subclass = re.search(r'^([^\s.]+)', lc_code)
        class_value = subclass.group(1)
    else: class_value = '999'

    return class_value

df['Class_value'] = df['CLASIFICACION(LC)'].apply(get_lc_subclass)
df['LC_MainClass'] = df['CLASIFICACION(LC)'].apply(lambda x: ''.join(filter(str.isalpha, x))[:2])

books_by_author = df['AUTOR'].value_counts()
books_by_publisher = df['EDITORIAL'].value_counts()
books_by_lc = df['LC_MainClass'].value_counts()
books_by_lc_detailed = df['Class_value'].value_counts()

with pd.ExcelWriter('enst_analysis.xlsx') as writer:
    books_by_lc.to_excel(writer, sheet_name='Class')
    books_by_lc_detailed.to_excel(writer, sheet_name='Class-Detail')
    books_by_author.to_excel(writer, sheet_name='Author')
    books_by_publisher.to_excel(writer, sheet_name='Publisher')