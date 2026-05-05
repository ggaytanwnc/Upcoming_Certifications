import pandas as pd
import openpyxl

def get_data(shared_folder,filename, sheet):
    df = pd.read_excel((shared_folder + "\\"+filename), sheet_name=sheet)
    #removiendo columnas y filas no necesarias
    df = df.iloc[10:,3:]

    #Conservando columnas por usar
    df = df.iloc[:, :8]

    #promoviendo primera fila como titulos de columnas
    df.columns = df.iloc[0]
    df = df[1:]

    #removiendo nulos
    df = df.dropna()

    #Obteniendo dias de diferencia entre fecha de expiracion y fecha actual
    df['Dias por expirar'] = (pd.to_datetime(df['FECHA DE RECERTIFICACION']) - pd.to_datetime('today')).dt.days

    #filtrar por dias por expirar menores a 100
    df = df[df['Dias por expirar'] < 100]

    #Removiendo columnas no necesarias
    df = df.drop(columns=['DIAS A VENCER','ALERTA', 'PARAMETRO DE CERTIFICACION'])

    #Dejando columnas fecha solo con formato fecha
    df['FECHA DE RECERTIFICACION'] = pd.to_datetime(df['FECHA DE RECERTIFICACION']).dt.date
    df['FECHA DE CERTIFICACION MM/DD/AAAA'] = pd.to_datetime(df['FECHA DE CERTIFICACION MM/DD/AAAA']).dt.date
    return df

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