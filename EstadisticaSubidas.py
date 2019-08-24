import datetime
import numpy as np
import pandas as pd
from matplotlib import pyplot as plt
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

rutvalues = sheet.values_get(ruts)["values"]; asistvalues = sheet.values_get(asistencia)["values"]; statusvalues = sheet.values_get(status)["values"]
seccionvalues = sheet.values_get(seccion)["values"]; fechavalues = sheet.values_get(fecha)["values"]

for i,a in enumerate(asistvalues):
	if "Presente" in a:
		print(rutvalues[i][0],seccionvalues[i][0],fechavalues[i][0])

#asistentes = 
