#importar pandas
import pandas as pd

#importar sqldf para manejar sql
from pandasql import sqldf

#leer el archivo y la hoja
df=pd.read_excel(open('Sistema de cobro 2-01-2018.xlsx','rb'), sheet_name='Asignacion')

pysqldf = lambda q: sqldf(q, globals())


#manejo de sql con dataframe
print(pysqldf("SELECT SUBCAMPANAPORCLIENTE as fase,sum(SALDO)as valor,count(*) as asignados  FROM df group by SUBCAMPANAPORCLIENTE order by valor desc;"))
print(pysqldf("""SELECT  CASE 
WHEN DIAS_MORA<=30 THEN 'DE 0 A 30 DIAS'
WHEN DIAS_MORA>30 and DIAS_MORA<=60 THEN 'DE 31 A 60 DIAS'
WHEN DIAS_MORA>60 and DIAS_MORA<=90 THEN 'DE 61 A 90 DIAS'
WHEN DIAS_MORA>90 and DIAS_MORA<=180 THEN 'DE 91 A 180 DIAS'
WHEN DIAS_MORA>180 and DIAS_MORA<=270 THEN 'DE 180 A 270 DIAS'
WHEN DIAS_MORA>270 and DIAS_MORA<=360 THEN 'DE 270 A 360 DIAS'
WHEN DIAS_MORA>360 and DIAS_MORA<=720 THEN 'DE 360 A 720 DIAS'
WHEN DIAS_MORA>720 and DIAS_MORA THEN 'MAYOR A 720 DIAS'
ELSE  'OTRO RANGO'
END  as altura_mora, sum(SALDO)as valor,count(*) as asignados  FROM df group by altura_mora order by valor desc;"""))

#estadisticos con dataframe
print(df[['SALDO','DIAS_MORA','CODIGO_CONDICION']].describe())