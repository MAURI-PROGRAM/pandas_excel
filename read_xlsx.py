#importar pandas
import pandas as pd

#leer el archivo y la hoja
df=pd.read_excel(open('archivo.xlsx','rb'), sheet_name='Sheet1')

#presentar lo que se leyo del .xlsx
print(df).
