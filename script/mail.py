from email.message import EmailMessage
import win32com.client as win32
from config import EMAIL_SENDER


def send_email(receiver, subject, body, attachment_path):
    #msg =  EmailMessage()
    #msg['From'] = sender
    #msg['To'] = receiver
    #msg['Subject'] = subject
    #msg.set_content(body)

    #obteniendo configuracion cuenta BOT
    sender = EMAIL_SENDER

    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.SentOnBehalfOfName = sender
        mail.To = receiver
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
    