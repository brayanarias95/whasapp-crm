from datetime import datetime
import pywhatkit
import time
import os
from openpyxl import load_workbook
from openpyxl import Workbook


wd = os.getcwd()
dirname=wd.replace('\\' , '/')

filesheet=dirname+"/prueba2.xlsx"
wb = load_workbook(filesheet)
hojas= wb.get_sheet_names()
print(hojas)
LIBRO=wb.get_sheet_by_name('Hoja1')
time.sleep(2)
#recorro el excel
for i in range(2,6):
    plac, B, C, D, propie, f, g, h,vigente, donde = LIBRO[f'A{i}:J{i}'][0]


    segundos=time.time()+60
    date=datetime.fromtimestamp(segundos)
    print(date.hour)
    print(date.second)
    try:
    #Enviando mensaje usando pywhatkiit
        pywhatkit. sendwhatmsg("+57"+str(g.value),"hola la motocicleta de placa "+str(plac.value)+" no ha hecho la revision te esperamos",date.hour,date.minute)
        print("Abriendo . .. Escanee Codig")
        print("Enviado exitosamente")
    except:
        print("No se pudo encontrar")

    
    time.sleep(1)
    
 





