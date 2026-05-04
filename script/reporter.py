import pandas as pd
import os

def export_report(path,df):
    try:
        df.to_excel(path, index=False)
        print(f"Report exported successfully to {path}")
    except Exception as e:
        print(f"Something wrong happened: {e}")