import pandas as pd
import numpy as np

# Specifikujte cestu k souboru Excel
excel_file_path = 'deal.xlsx'

# Načtení dat ze souboru Excel do DataFrame
df = pd.read_excel(excel_file_path)
df['ID'] = np.arange(1, len(df) + 1)




# Tisk prvních pěti řádků DataFrame
print(df.head())