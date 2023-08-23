import pandas as pd
import numpy as np

# Specifikujte cestu k souboru Excel
excel_file_path = 'deal.xlsx'
excel_file_path1 = 'crazy.xlsx'

# Načtení dat ze souboru Excel do DataFrame

df = pd.read_excel(excel_file_path)
df['ID'] = np.arange(1, len(df) + 1)

df1 = pd.read_excel(excel_file_path1)
df1['ID'] = np.arange(1, len(df1) + 1)


# Funkce pro přidání K na začátek u deal položek
def pridat_K_k_deal(item, prefix):
    return prefix + item

# Přidání znaku na začátek položek v prvním sloupci
prefix = 'K'
df['code'] = df['code'].apply(pridat_K_k_deal, args=(prefix,))

#vyfiltrovat řádky s hodnotou skladu >0
mask = df['stock'] != 0
df = df[mask]

#zboží se zásobou větší než 3 bude automaticky poníženo na 3
def maximalni_zasoba_deal(value):
    return min(value, 3)

df['stock'] = df['stock'].apply(maximalni_zasoba_deal)



# Tisk řádků DataFrame
print(df.head())

# Pro Crazy které začínají W a mají stock=0, místo 0 tam pak dosadíš zásobu z dealu a přejmenuješ na KW. Pokud se v deal nenachází, tak s tím nedělám nic. 