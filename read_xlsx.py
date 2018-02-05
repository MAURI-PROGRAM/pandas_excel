#importar pandas
import pandas as pd

#importar sqldf para manejar sql
from pandasql import sqldf

#leer el archivo y la hoja
df=pd.read_excel(open('archivo.xlsx','rb'), sheet_name='Asignacion')

pysqldf = lambda q: sqldf(q, globals())


#manejo de sql con dataframe y pandasql
presentacion1=pysqldf("SELECT SUBCAMPANAPORCLIENTE as fase,sum(SALDO)as valor,round(sum(SALDO)/(select sum(SALDO)FROM df),2) as porcentaje_valor,count(*) as asignados  FROM df group by SUBCAMPANAPORCLIENTE order by valor desc;")
print(presentacion1)

presentacion2=pysqldf("SELECT count(DISTINCT IDENTIFICACION_DEUDOR) as No_CEDULAS,count(*)as No_MULTAS ,SUM(SALDO) as MONTO_TOTAL   FROM df ;")
print(presentacion2)

presentacion3=pysqldf("""SELECT  CASE 
WHEN DIAS_MORA<=30 THEN 'DE 0 A 30 DIAS'
WHEN DIAS_MORA>30 and DIAS_MORA<=60 THEN 'DE 31 A 60 DIAS'
WHEN DIAS_MORA>60 and DIAS_MORA<=90 THEN 'DE 61 A 90 DIAS'
WHEN DIAS_MORA>90 and DIAS_MORA<=180 THEN 'DE 91 A 180 DIAS'
WHEN DIAS_MORA>180 and DIAS_MORA<=270 THEN 'DE 180 A 270 DIAS'
WHEN DIAS_MORA>270 and DIAS_MORA<=360 THEN 'DE 270 A 360 DIAS'
WHEN DIAS_MORA>360 and DIAS_MORA<=720 THEN 'DE 360 A 720 DIAS'
WHEN DIAS_MORA>720 and DIAS_MORA THEN 'MAYOR A 720 DIAS'
ELSE  'OTRO RANGO'
END  as altura_mora, sum(SALDO)as valor,count(*) as asignados  FROM df group by altura_mora order by valor desc;""")
print(presentacion3)

#estadisticos con dataframe
presentacion4=df[['SALDO','DIAS_MORA']].describe()
print(presentacion4)

writer = pd.ExcelWriter('Detalle.xlsx', engine='xlsxwriter')
workbook = writer.book 

format1 = workbook.add_format()
format1.set_center_across()
format2 = workbook.add_format({'num_format': '0%'})
format2.set_center_across()
format3 = workbook.add_format({'num_format': '#.##0.00'})
format3.set_center_across()

presentacion2.to_excel(writer, sheet_name='Asignacion')
worksheet2 = writer.sheets['Asignacion']
worksheet2.set_column('B:B', 20, format1) 
worksheet2.set_column('C:C', 18, format1)
worksheet2.set_column('D:D', 18, format1) 


presentacion1.to_excel(writer, sheet_name='ASIG_FASES')
worksheet = writer.sheets['ASIG_FASES']
worksheet.set_column('B:B', 20, format1) 
worksheet.set_column('C:C', 18, format1) 
worksheet.set_column('D:D', 16, format2)
worksheet.set_column('E:E', 16, format1)


presentacion3.to_excel(writer, sheet_name='ASIG_MORA')
worksheet = writer.sheets['ASIG_MORA']
worksheet.set_column('B:B', 20, format1) 
worksheet.set_column('C:C', 18, format1) 
worksheet.set_column('D:D', 16, format1)

presentacion4.to_excel(writer, sheet_name='ESTADISTICOS')
worksheet = writer.sheets['ESTADISTICOS']
worksheet.set_column('B:B', 20, format1) 
worksheet.set_column('C:C', 18, format1) 

writer.save()