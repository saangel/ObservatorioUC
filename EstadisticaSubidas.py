import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials

ruts = 'Calendario!E1:E'
asistencia = 'Calendario!J1:J'
status = 'Calendario!I1:I'
seccion = 'Calendario!F1:F'
fecha = 'Calendario!G1:G'

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('EstadisticasObsUC.json',scope)
client = gspread.authorize(creds)
sheet = client.open("Calendario visitas Obs UC - Astronomía AST0111 - 2019.2")

hojaAsistencias = sheet.worksheet("Asistencias")

rutvalues = sheet.values_get(ruts)["values"]; asistvalues = sheet.values_get(asistencia)["values"]
seccionvalues = sheet.values_get(seccion)["values"]; fechavalues = sheet.values_get(fecha)["values"]

asistencias = []
for i,a in enumerate(asistvalues):
	if "Presente" in a:
		asistencias.append([rutvalues[i][0],seccionvalues[i][0],fechavalues[i][0]])

for i,a in enumerate(asistencias):
	date=datetime.datetime.strptime(a[2],"%d/%m/%Y")
	hojaAsistencias.update_cell(41+i,1,a[0])
	hojaAsistencias.update_cell(41+i,2,a[1])
	hojaAsistencias.update_cell(41+i,4,date.strftime("%W"))
	hojaAsistencias.update_cell(41+i,3,date.strftime("%m"))
#end

'''
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
'''
