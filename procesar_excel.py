"""
Módulo para procesar archivos Excel relacionados con vencimientos de licencias.
Incluye funciones para leer, validar y actualizar datos en el archivo Excel.
"""

import pandas as pd
from datetime import datetime, timedelta

# Constantes
ARCHIVO_EXCEL = "VencimientosClientesHB.xlsx"
COLUMNAS_REQUERIDAS = [
    "#", "Tipo", "Serial", "Empresa", "Creacion", "UltimaCon", "Vencimiento",
    "Distribuidor", "CodProd", "Producto", "Contacto", "email", "Status",
    "NotifSend", "FechaSend"
]

def leer_excel(archivo):
    """
    Lee un archivo Excel y devuelve un DataFrame de pandas.
    
    Args:
        archivo (str): Ruta al archivo Excel.
    
    Returns:
        pd.DataFrame: DataFrame con los datos del archivo.
    """
    try:
        df_local = pd.read_excel(archivo)
        print("Archivo Excel leído correctamente.")
        return df_local
    except FileNotFoundError as e:
        print(f"Archivo no encontrado: {e}")
        return None
    except pd.errors.EmptyDataError as e:
        print(f"El archivo está vacío o no contiene datos válidos: {e}")
        return None
    except Exception as e:
        print(f"Error inesperado al leer el archivo Excel: {e}")
        return None

def verificar_archivo_vacio(df_local):
    """
    Verifica si el DataFrame está vacío.
    
    Args:
        df_local (pd.DataFrame): DataFrame a verificar.
    
    Returns:
        bool: True si está vacío, False en caso contrario.
    """
    if df_local.empty:
        print("El archivo está vacío.")
        return True
    else:
        print("El archivo contiene datos.")
        return False

def contar_registros(df_local):
    """
    Cuenta los registros en el DataFrame.
    
    Args:
        df_local (pd.DataFrame): DataFrame a verificar.
    
    Returns:
        int: Número de registros.
    """
    num_registros = len(df_local)
    print(f"El archivo contiene {num_registros} registros.")
    return num_registros

def validar_columnas(df_local, columnas_requeridas):
    """
    Valida que todas las columnas requeridas estén presentes en el DataFrame.
    
    Args:
        df_local (pd.DataFrame): DataFrame a verificar.
        columnas_requeridas (list): Lista de columnas requeridas.
    
    Returns:
        bool: True si todas las columnas están presentes, False en caso contrario.
    """
    columnas_faltantes = [col for col in columnas_requeridas if col not in df_local.columns]
    if columnas_faltantes:
        print(f"Error: Faltan las siguientes columnas en el archivo: {', '.join(columnas_faltantes)}")
        return False
    else:
        print("El archivo contiene todas las columnas requeridas.")
        return True

def verificar_vencimiento_vacio(df_local):
    """
    Verifica si la columna 'Vencimiento' está completamente vacía.
    
    Args:
        df_local (pd.DataFrame): DataFrame a verificar.
    
    Returns:
        bool: True si está vacía, False en caso contrario.
    """
    if df_local["Vencimiento"].isnull().all():
        print("La columna 'Vencimiento' está completamente vacía.")
        return True
    else:
        print("La columna 'Vencimiento' contiene datos.")
        return False

def calcular_fecha_vencimiento(fecha_creacion):
    """
    Calcula la próxima fecha de vencimiento sumando un año a la fecha de creación.
    
    Args:
        fecha_creacion (str): Fecha de creación en formato "YYYY-MM-DD".
    
    Returns:
        str: Fecha de vencimiento en formato "YYYY-MM-DD".
    """
    fecha_creacion = datetime.strptime(fecha_creacion, "%Y-%m-%d")
    fecha_vencimiento = fecha_creacion + timedelta(days=365)
    return fecha_vencimiento.strftime("%Y-%m-%d")

def corregir_correos(df_local):
    """
    Convierte todas las direcciones de correo electrónico a minúsculas.
    
    Args:
        df_local (pd.DataFrame): DataFrame a corregir.
    
    Returns:
        pd.DataFrame: DataFrame con correos corregidos.
    """
    correos_mayusculas = df_local[df_local["email"].str.contains(r'[A-Z]', na=False)]
    if not correos_mayusculas.empty:
        print("Se encontraron correos en mayúsculas. Corrigiendo...")
        df_local["email"] = df_local["email"].str.lower()
        print("Correos convertidos a minúsculas.")
    else:
        print("Todos los correos ya están en minúsculas.")
    return df_local

if __name__ == "__main__":
    # Leer el archivo Excel
    df = leer_excel(ARCHIVO_EXCEL)

    if df is not None:
        # Verificar si el archivo está vacío
        if verificar_archivo_vacio(df):
            exit()

        # Contar los registros
        contar_registros(df)

        # Validar las columnas requeridas
        if not validar_columnas(df, COLUMNAS_REQUERIDAS):
            exit()

        # Mostrar las primeras filas del archivo
        print("\nPrimeras filas del archivo:")
        print(df.head())

        # Verificar y calcular las fechas de vencimiento si están vacías
        if verificar_vencimiento_vacio(df):
            df["Vencimiento"] = df["Creacion"].apply(calcular_fecha_vencimiento)
            print("Fechas de vencimiento calculadas y actualizadas.")

        # Corregir las direcciones de correo electrónico
        df = corregir_correos(df)

        # Mostrar los datos actualizados
        print("\nDatos actualizados:")
        print(df[["#", "Empresa", "Creacion", "Vencimiento", "email"]])

        # Guardar los cambios en el archivo Excel
        try:
            df.to_excel(ARCHIVO_EXCEL, index=False)
            print(f"El archivo '{ARCHIVO_EXCEL}' ha sido actualizado correctamente.")
        except Exception as e:
            print(f"Error al guardar el archivo Excel: {e}")