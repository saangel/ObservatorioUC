import datetime
import numpy as np
import pandas as pd
from matplotlib import pyplot as plt

planilla=pd.ExcelFile("Calendario.xlsx")
calendario=planilla.parse("Calendario",na_filter=False)
listas=planilla.parse("ListasCursos",na_filter=False)

#cada una de las siguientes columnas tiene un "header" de aprox. 10 celdas
asistentes=calendario[calendario.keys()[1]]  
nombres=calendario["Unnamed: 2"]
fechas=calendario["Unnamed: 5"][9:].str.strip()
status=calendario["Unnamed: 7"].str.lower()
asistencia=calendario["Unnamed: 8"]
rut=calendario["Unnamed: 3"].str.replace(".","")
fechas=pd.to_datetime(fechas[fechas!="Fecha"],format="%d/%m/%y")

#"header" de 2 celdas
listas1=listas["AST0111-1"]#.fillna("")
listas2=listas["AST0111-2"]#.fillna("")
listas3=listas["AST0111-3"]#.fillna("")
listaeng=listas["AST0112-1"]#.fillna("")
listaetu=listas["Curso Rolando"]#.fillna("")
secciones=["S1","S2","S3","ENG","ETU"]
meses=[3,4,5,6]

listass2=[listas1,listas2,listas3,listaeng,listaetu]
listas=[]
for l in listass2:
	filt = l != ""
	l = l[filt]
	listas.append(l)
nsec=len(secciones)
total_listas=[len(l) for l in listas]
sum_asistencia=[]
meses=[]
semanas=[]
seccion=[]
nasistente=[]
rasistente=[]
for k,l in enumerate(listas):
	#print(secciones[k])
	sum_asistencia.append(0)
	for i in l[1:]:
		if i:
		#	print(i)
			asist=asistentes[asistentes.str.lower()==i.split(",")[0].lower()]
			rasist=rut[rut==i.replace(".","")]
			#print(asist+rasist)
			if len(asist):
				idx= asist.index
				
				#print(asist,idx)
			elif  len(rasist):
				idx= rasist.index
				#print(rasist,idx)
			
			if len(asist+rasist):
				for ii in idx:
					#print(asistentes[ii],fechas[ii].date().isoformat(),status[ii],asistencia[ii])
					if status[ii].lower() == "confirmada" and asistencia[ii].lower() == "presente":
					#	print(asistentes[ii],fechas[ii].weekofyear,fechas[ii].month)
						nasistente.append(nombres[ii]+" "+asistentes[ii])
						rasistente.append(rut[ii])
						seccion.append(secciones[k])
						sum_asistencia[k]+=1
						meses.append(fechas[ii].month)
						semanas.append(fechas[ii].weekofyear)

df=pd.DataFrame()
df["Nombre"]=nasistente
df["Seccion"]=seccion
df["Mes"]=meses
df["Semana"]=semanas
df["Rut"]=rasistente

df.to_excel("asistencias.xlsx")
