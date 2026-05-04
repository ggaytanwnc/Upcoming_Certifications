import pandas as pd
import openpyxl

def get_data(path, sheet_number):
    try:
        df = pd.read_excel(path, sheet_number)
        return df
    except Exception as e:
        print(f"Something wrong happened: {e}")
        return None