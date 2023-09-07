# RecuperationDonneesQuestionnaire

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.



import pandas as pd
import openpyxl

engine ='xlsxwriter'

df = pd.DataFrame([[11, 21, 31], [12, 22, 32], [31, 32, 33]], index=['one', 'two', 'three'], columns=['a', 'b', 'c'])
print(df)
df.to_excel('/Users/evelyne/PycharmProjects/ATinput/pandas_to_excel.xlsx', sheet_name='new_sheet_name')
