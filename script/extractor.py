import pandas as pd
import openpyxl
import numpy as np
from math import nan

def get_data(shared_folder,filename, sheet):
    df_raw = pd.read_excel((shared_folder + "\\"+filename), sheet_name=sheet, header=None)

    header_row = find_header_row(df_raw, ['PID', 'NOMBRE', 'CERTIFICACION'])

    #df = pd.read_excel((shared_folder + "\\"+filename), sheet_name=sheet, header=header_row)
    
    #encontrando columna para iniciar
    header_values = df_raw.iloc[header_row]
    start_col = header_values[
        header_values.astype(str).str.contains('PID', case=False, na=False)
    ].index[0]

    #Dataframe limpio
    df = df_raw.iloc[header_row:, start_col:]

    #removiendo columnas y filas no necesarias
    #df = df.iloc[0:,3:]

    #Conservando columnas por usar
    #df = df.iloc[:, :8]

    #promoviendo primera fila como titulos de columnas
    df.columns = df.iloc[0]
    df = df[1:]

    # convertir espacios vacíos en NaN
    df = df.replace(r'^\s*$', pd.NA, regex=True)

    #removiendo nulos
    empty_col_index = None

    for i, col in enumerate(df.columns):
        if pd.isna(col):
            empty_col_index = i
            break


    if empty_col_index is not None:
        df = df.iloc[:, :empty_col_index]


    #Obteniendo dias de diferencia entre fecha de expiracion y fecha actual
    df['Dias por expirar'] = (pd.to_datetime(df['FECHA DE RECERTIFICACION']) - pd.to_datetime('today')).dt.days + 1

    #filtrar por dias por expirar menores a 100
    df = df[df['Dias por expirar'] < 100]

    #Removiendo columnas no necesarias
    df = df.drop(columns=['DIAS A VENCER','ALERTA', 'PARAMETRO DE CERTIFICACION'])

    #Conservando solo columnas necesarias


    #Dejando columnas fecha solo con formato fecha
    df['FECHA DE RECERTIFICACION'] = pd.to_datetime(df['FECHA DE RECERTIFICACION']).dt.date
    df['FECHA DE CERTIFICACION MM/DD/AAAA'] = pd.to_datetime(df['FECHA DE CERTIFICACION MM/DD/AAAA']).dt.date

    #removiendo nulos de primera columna
    df = df.dropna(subset=[df.columns[0]])
    df = df.dropna(subset=[df.columns[3]])
    return df


def find_header_row(df_raw, required_headers):
    for idx, row in df_raw.iterrows():
        row_values = row.astype(str).tolist()
        if all(header in row_values for header in required_headers):
            return idx
    
    raise Exception("Header row not found")

"""
#Version Anterior
def get_data(path, sheet_number):
    try:
        df = pd.read_excel(path, sheet_number)
        return df
    except Exception as e:
        print(f"Something wrong happened: {e}")
        return None
"""