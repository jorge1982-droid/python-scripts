# -*- coding: utf-8 -*-
"""
Created on Wed Apr 14 19:14:14 2021

@author: Lyde
"""

import pandas as pd
import pyodbc 
import pyfiglet
import  time


ascii_banner = pyfiglet.figlet_format("Lyde Reports ")
print(ascii_banner)


archivo=input("Ingrese Nombre del Reporte :  ")

print("\n")

numero1=input("ingrese numero de cliente: ")
numero2=input("ingrese numero provedor: ")
fechaI=input("Ingresa fecha Inicial: ")
fechaF=input("Ingresa Fecha Final: ")


#conexion a sql
start = time.time()
server = '198.38.94.40' 
database = 'wms' 
username = 'analisis01' 
PWD = 'Analisis_01' 
cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ PWD)
cursor = cnxn.cursor()



if numero2 =='':
    numero2= 'null'
    

    
    
###Consultas SQL

sql="select * from wms.dbo.ufn_Reporte_ELEKTRA_Trazabilidad("+numero1+","+numero2+")"
datos=pd.read_sql_query(sql,con=cnxn)
df1=pd.DataFrame(datos)
df1['SKU']=df1['SKU'].astype(int)




sql2=" select * from wms.dbo.[ufn_Reporte_Ordenes_Solicitudes_Salida]( "+numero1+", "+numero2+", '"+fechaI+"', '"+fechaF+"')"
datos2=pd.read_sql_query(sql2,con=cnxn)
df2=pd.DataFrame(datos2)
df2['NumeroTienda']=df2['NumeroTienda'].astype(int)
df2['SKU']=df2['SKU'].astype(int)


sql3=" select * from wms.dbo.uf_Trazabilidad("+numero1+")"
datos3=pd.read_sql_query(sql3,con=cnxn)
df3=pd.DataFrame(datos3)


sql4 = "select TA_Folio, ISNULL(TA_FolioRemision,'') as FolioRemision ,"
sql4 += "ISNULL(TA_FolioRuta,'') as FolioRuta ,ISNULL(TAF_FolioEntrada,'') as FolioEntrada ,"
sql4 += "ISNULL(TAF_FolioCargo,'') as FolioCargo from TransferenciaAlmacen_FoliosEKT f, TransferenciaAlmacen t where t.TA_ID = f.TA_ID and t.TA_ID > 1"

datos4=pd.read_sql_query(sql4,con=cnxn)
df4=pd.DataFrame(datos4)
df4['FolioRemision']=df4['FolioRemision'].astype(int)




writer = pd.ExcelWriter(""+archivo+".xlsx")
df1.to_excel(writer, sheet_name="Inventarios", index=False)
df2.to_excel(writer,sheet_name="Dos " ,index=False)
df3.to_excel(writer,sheet_name="Resumen Ordenes",index=False)
df4.to_excel(writer,sheet_name="Documentos",index=False)


 

ascii_banner = pyfiglet.figlet_format("   !!!Reporte Terminado !!!")
print(ascii_banner)

end = time.time()

print("Tiempo de ejecucion",end - start,"segundos")
writer.save()
writer.close()
