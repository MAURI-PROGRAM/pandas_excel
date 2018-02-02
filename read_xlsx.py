#importar pandas
import pandas as pd

#importar sqldf para manejar sql
from pandasql import sqldf

#leer el archivo y la hoja
df=pd.read_excel(open('archivo.xlsx','rb'), sheet_name='Sheet1')

pysqldf = lambda q: sqldf(q, globals())

#presentar lo que se leyo del .xlsx
print(df)

#manejo de sql con dataframe
print (pysqldf("SELECT id,sum(valores) as saldo_tl,count(*) as n_multas,min(dias) as min_dia,max(dias) as max_dia  FROM df group by id;").head())

