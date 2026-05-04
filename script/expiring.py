#Librerias Import
import pandas as pd
import numpy as np
import smtplib
from email.message import EmailMessage
import win32com.client as win32
import os

#Data Handling
base_dir = os.path.dirname(os.path.abspath(__file__))
read_path = os.path.join(base_dir, '..', 'data', 'Copy of 10.-SMD.xlsm')
#df_certifications = pd.read_excel('../data/Copy of 10.-SMD.xlsm',sheet_name='OP 1 C')
df_certifications = pd.read_excel(read_path, sheet_name='OP 1 C')

#removiendo columnas y filas no necesarias
df = df_certifications.iloc[10:,3:]

#Conservando columnas por usar
df = df.iloc[:, :8]

#promoviendo primera fila como titulos de columnas
df.columns = df.iloc[0]
df = df[1:]

#removiendo nulos
df = df.dropna()

print("here")
df.to_excel('test.xlsx', index=False)

#Obteniendo dias de diferencia entre fecha de expiracion y fecha actual
df['Dias por expirar'] = (pd.to_datetime(df['FECHA DE RECERTIFICACION']) - pd.to_datetime('today')).dt.days

#filtrar por dias por expirar menores a 100
df = df[df['Dias por expirar'] < 100]

#Removiendo columnas no necesarias
df = df.drop(columns=['DIAS A VENCER','ALERTA', 'PARAMETRO DE CERTIFICACION'])

#Dejando columnas fecha solo con formato fecha
df['FECHA DE RECERTIFICACION'] = pd.to_datetime(df['FECHA DE RECERTIFICACION']).dt.date
df['FECHA DE CERTIFICACION MM/DD/AAAA'] = pd.to_datetime(df['FECHA DE CERTIFICACION MM/DD/AAAA']).dt.date

extract_path = os.path.join(base_dir, '..', 'data', 'reporte.xlsx')
df.to_excel(extract_path, index=False)
ruta_archivo = os.path.abspath(extract_path)

#Email
sender = "sender@email.com"
receiver = "receiver.com"
pwd = 'paswword'

msg  = EmailMessage()
msg['From'] = sender
msg['To'] = receiver   
msg['Subject'] = "Certificaciones por expirar en menos de 100 dias"
msg.set_content("Hola, adjunto el reporte de certificaciones por expirar en menos de 100 dias")


outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.SentOnBehalfOfName = sender
mail.To = receiver
mail.Subject = "Certificaciones por expirar"
mail.Body = "Buen dia, se comparte el reporte de certificaciones por expirar en menos de 100 dias \nFavor de incluir en programa semanal \nSaludos!"
mail.attachments.Add(ruta_archivo)
mail.Send()