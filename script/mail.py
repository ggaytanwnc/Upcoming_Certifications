from email.message import EmailMessage
import win32com.client as win32
from config import EMAIL_SENDER, df_Mail

#Obteniendo listado de mail
list_to = ";".join(df_Mail["To"].dropna().astype(str))
list_cc = ";".join(df_Mail["CC"].dropna().astype(str))

def send_email(receiver, cc,  subject, body, attachment_path):

    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.SentOnBehalfOfName = EMAIL_SENDER
        mail.To = receiver
        mail.CC = cc
        mail.Subject = subject
        mail.Body = body
        mail.Attachments.Add(attachment_path)
        mail.Send()
    except Exception as e:
        print(f"Something wrong happened: {e}")

def send_error_notification(error_message):
    sender = EMAIL_SENDER
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.SentOnBehalfOfName = EMAIL_SENDER
        mail.To = "gilberto.gaytan@wnc.com.tw"
        mail.Subject = "Error"
        mail.Body = str(error_message)
        mail.Send()
    except Exception as e:
        print(f"Something wrong happened: {e}")
    