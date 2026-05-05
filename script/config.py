import os
import pandas as pd
from datetime import datetime

EMAIL_SENDER = 'p_m1dmis_rpa01@wnc.com.tw'
PWD = 'Wnewebmx2026.'
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PATH_CONFIG_FILE = os.path.join(BASE_DIR, 'Config\Config.xlsx')
FILENAME_REPORT = f"Training_Report_{datetime.now().strftime('%Y%m%d')}.xlsx"
PATH_REPORT_FILE = os.path.join(BASE_DIR, 'reports',FILENAME_REPORT)

#Data from Config File
df_config = pd.read_excel(PATH_CONFIG_FILE, sheet_name="Config")
df_data = pd.read_excel(PATH_CONFIG_FILE, sheet_name="Data")
df_Mail = pd.read_excel(PATH_CONFIG_FILE, sheet_name="Mail")