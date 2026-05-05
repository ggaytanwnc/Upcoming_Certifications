from config import BASE_DIR, PATH_CONFIG_FILE, df_config, df_data, PATH_REPORT_FILE, df_Mail
from extractor import get_data
from transformer import transform_data
from reporter import export_report
from mail import send_email, send_error_notification
import os
import warnings
import pandas as pd


warnings.filterwarnings("ignore")

try:
    #Reading Config File (Path of Data)
    shared_folder = df_config.iloc[0,1]

    #Extract Data and Transform Data
    report = []
    for row in df_data.itertuples():
        filename = row.Filename
        sheet = row.Sheet

        df = get_data(shared_folder, filename, sheet)
        report.append(df)

    df_training_report = pd.concat(report, ignore_index=True)

    
    #Export Data
    export_report(PATH_REPORT_FILE, df_training_report)

    #Sending a email
    #Obteniendo listado de mail
    list_to = ";".join(df_Mail["To"].dropna().astype(str))
    list_cc = ";".join(df_Mail["CC"].dropna().astype(str))
    subject = 'Certificaciones por expirar'
    body = 'Adjunto el reporte de certificaciones por expirar en los próximos 100 días.'

    send_email(list_to, list_cc, subject, body, PATH_REPORT_FILE)
    print("Process completed successfully.")

except Exception as e:
    print(e)
    send_error_notification(e)