from config import BASE_DIR
from extractor import get_data
from transformer import transform_data
from reporter import export_report
from mail import send_email, send_error_notification
import os
import warnings

warnings.filterwarnings("ignore")

try:
    source_report = 'Copy of 10.-SMD.xlsm'
    path_report = os.path.join(BASE_DIR, '..', 'data', source_report)
    sheet_number = "OP 1 C"
    ##print(path_report)

    #Extract Data
    df_training_report = get_data(path_report, sheet_number)

    #Transform Data
    df_training_report = transform_data(df_training_report)

    #Export Data
    report_name = 'training_cert_expiring.xlsx'
    path_export = os.path.join(BASE_DIR, '..', 'reports', report_name)

    export_report(path_export, df_training_report),1


    #Sending a email
    receiver = 'gilberto.gaytan@wnc.com.tw'
    subject = 'Certificaciones por expirar'
    body = 'Adjunto el reporte de certificaciones por expirar en los próximos 100 días.'
    attachment_path = path_export
    send_email(receiver, subject, body, attachment_path)
    print("Process completed successfully.")

except Exception as e:
    send_error_notification(e)