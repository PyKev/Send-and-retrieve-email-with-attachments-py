import win32com.client as win32
import pandas as pd
import os
import openpyxl
from datetime import date

def enviar_recordatorios():
    df = pd.read_excel('Datos.xlsx')

    for i in range(len(df)):
        nombre = df.iloc[i]['Responsable']
        correo= df.iloc[i]['Correo']
        mensaje = df.iloc[i]['Mensaje']
        plazo = df.iloc[i]['Plazo']
        plazo = date(plazo.year, plazo.month, plazo.day)
        hoy = date.today()
        dias = plazo-hoy
        dias = str(dias)
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = correo
        mail.Subject = 'Recordatorio'
        mail.Body = nombre + " te quedan " + dias[0:dias.find("days")].strip() + " días para enviar tu trabajo " + mensaje
        mail.Send()

def trae_correos_y_adjuntos():
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Ultimos 3 correos
    inbox = outlook.GetDefaultFolder(6)
    emails = inbox.Items
    emails.Sort("[ReceivedTime]", True)

    wb = openpyxl.Workbook()
    ws = wb.active
    cabeceras = ["Remitente", "Fecha", "Asunto"]

    for i in range(0, 3):
        correo = emails[i]
        remitente = correo.SenderEmailAddress
        fecha = correo.ReceivedTime.strftime("%d/%m/%Y %I:%M:%S %p")
        asunto = correo.Subject

        lista = [remitente, fecha, asunto]
        contador = 0
        for archivo in correo.Attachments:
            if archivo.FileName.lower().endswith((".pdf", ".doc", ".docx")):
                temp_file = os.path.join(os.getcwd(), archivo.FileName)
                archivo.SaveAsFile(temp_file)
                hyperlink = f'=HYPERLINK("{temp_file}", "{archivo.FileName}")'
                try:
                    cabeceras.append(f"Archivo {contador}")
                except:
                    pass
                lista.append(hyperlink)
            contador += 1
        try:
            ws.append(cabeceras)
            del cabeceras
        except:
            pass
        ws.append(lista)

    wb.save("correos.xlsx")

enviar_recordatorios()
trae_correos_y_adjuntos()

