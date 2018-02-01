import pandas as pd
import xlrd

df=pd.read_excel(open('hoja.xlsx','rb'), sheet_name='Sheet1')
print(df.columns)
