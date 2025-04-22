"""
Módulo para procesar archivos Excel relacionados con vencimientos de licencias.
Incluye funciones para leer, validar y actualizar datos en el archivo Excel.
"""
# Imports estándar
import datetime
from datetime import datetime, timedelta

# Imports de terceros
import pandas as pd

# Constantes
ARCHIVO_EXCEL = "VencimientosClientesHB.xlsx"
COLUMNAS_REQUERIDAS = [
    "#", "Tipo", "Serial", "Empresa", "Creacion", "UltimaCon", "Vencimiento",
    "Distribuidor", "CodProd", "Producto", "Contacto", "email", "Status",
    "NotifSend", "FechaSend", "FechaActual", "HoraActual", "FechaRonavada", "StatusLicencia"
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
        df_local = pd.read_excel(archivo, dtype={"Serial": str})
        print("Archivo Excel leído correctamente.")
        return df_local
    except FileNotFoundError as e:  # Captura errores de archivo no encontrado
        print(f"Archivo no encontrado: {e}")
        return None
    except pd.errors.EmptyDataError as e:  # Captura errores de archivo vacío
        print(f"El archivo está vacío o no contiene datos válidos: {e}")
        return None
    except PermissionError as e:  # Captura errores de permisos
        print(f"No se tienen permisos para leer el archivo: {e}")
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
        print(
           f"Error: Faltan las siguientes columnas en el archivo: {', '.join(columnas_faltantes)}"
           )
        return False
    else:
        print("El archivo contiene todas las columnas requeridas.")
        return True

def actualizar_fecha_y_hora_actual(df_local):
    """
    Actualiza las columnas 'FechaActual' y 'HoraActual' con la fecha y hora actuales.
    
    Args:
        df_local (pd.DataFrame): DataFrame a verificar.
    
    Returns:
        pd.DataFrame: DataFrame con las columnas 'FechaActual' y 'HoraActual' actualizadas.
    """
    print("Actualizando las columnas 'FechaActual' y 'HoraActual'...")

    # Obtener la fecha y hora actuales
    fecha_actual = datetime.now().strftime("%Y-%m-%d")
    hora_actual = datetime.now().strftime("%H:%M:%S")

    # Sobrescribir las columnas 'FechaActual' y 'HoraActual'
    df_local["FechaActual"] = fecha_actual
    df_local["HoraActual"] = hora_actual

    print("Las columnas 'FechaActual' y 'HoraActual' han sido actualizadas.")
    return df_local

def limpiar_contactos(df_local):
    """
    Limpia los valores de la columna 'Contacto' eliminando los caracteres '**' al inicio y al final.
    
    Args:
        df_local (pd.DataFrame): DataFrame a limpiar.
    
    Returns:
        pd.DataFrame: DataFrame con la columna 'Contacto' limpia.
    """
    print("Limpiando la columna 'Contacto'...")
    if "Contacto" in df_local.columns:
        # Eliminar los caracteres '**' al inicio y al final usando .str.strip()
        df_local["Contacto"] = df_local["Contacto"].str.strip("*")
        print("La columna 'Contacto' ha sido limpiada.")
    else:
        print("La columna 'Contacto' no existe en el archivo.")
    return df_local

def ajustar_formato_empresa(df_local):
    """
    Ajusta el formato de la columna 'Empresa' para que el contenido esté justificado a la izquierda.
    
    Args:
        df_local (pd.DataFrame): DataFrame a ajustar.
    
    Returns:
        pd.DataFrame: DataFrame con la columna 'Empresa' ajustada.
    """
    print("Ajustando el formato de la columna 'Empresa'...")
    if "Empresa" in df_local.columns:
        # Asegurar que los valores sean cadenas y justificar a la izquierda
        df_local["Empresa"] = df_local["Empresa"].astype(str).str.ljust(50)
        print("La columna 'Empresa' ha sido ajustada.")
    else:
        print("La columna 'Empresa' no existe en el archivo.")
    return df_local
def calcular_fecha_vencimiento(fecha_creacion, fecha_renovada=None):
    """
    Calcula la fecha de vencimiento sumando 365 días a la fecha de creación.
    Args:
        fecha_creacion (str): Fecha de creación en formato "DD/MM/YYYY" o "YYYY-MM-DD".
        fecha_renovada (str, optional): Fecha de renovación en formato "DD/MM/YYYY" o "YYYY-MM-DD". Defaults to None.
    Returns:
        str: Fecha de vencimiento en formato "YYYY-MM-DD".
    """
    # Intentar parsear la fecha en varios formatos posibles
    try:
        fecha_creacion_dt = datetime.strptime(fecha_creacion, "%d/%m/%Y")
    except ValueError:
        try:
            fecha_creacion_dt = datetime.strptime(fecha_creacion, "%Y-%m-%d")
        except ValueError as exc:
            raise ValueError(
                f"Formato de fecha inválido: {fecha_creacion}."
                f"Se esperaba 'DD/MM/YYYY' o 'YYYY-MM-DD'."
            ) from exc
    
    # Calcular la fecha de vencimiento sumando 365 días a la fecha de creación
    fecha_vencimiento = fecha_creacion_dt + timedelta(days=365)
    
    # Si la fecha de vencimiento ya pasó y hay una fecha de renovación, sumar otro año completo
    if fecha_renovada:
        try:
            fecha_renovada_dt = datetime.strptime(fecha_renovada, "%d/%m/%Y")
        except ValueError:
            try:
                fecha_renovada_dt = datetime.strptime(fecha_renovada, "%Y-%m-%d")
            except ValueError as exc:
                raise ValueError(
                    f"Formato de fecha inválido: {fecha_renovada}."
                    f"Se esperaba 'DD/MM/YYYY' o 'YYYY-MM-DD'."
                ) from exc
        
        # Si la fecha de vencimiento ya pasó, sumar otro año completo
        hoy = datetime.now()
        if fecha_vencimiento < hoy:
            fecha_vencimiento += timedelta(days=365)
    
    return fecha_vencimiento.strftime("%Y-%m-%d")
def calcular_vencimiento(df_local):
    """
    Calcula la fecha de vencimiento para todos los registros basándose en las columnas 'Creacion' y 'FechaRenovada'.
    Las fechas de vencimiento se actualizan al año en curso o al próximo año, dependiendo de las condiciones.
    
    Args:
        df_local (pd.DataFrame): DataFrame a procesar.
    
    Returns:
        pd.DataFrame: DataFrame con las fechas de vencimiento actualizadas.
    """
    print("Calculando fechas de vencimiento para todos los registros...")
    
    # Verificar que las columnas 'Creacion' y 'FechaRenovada' existan
    if "Creacion" not in df_local.columns:
        print("La columna 'Creacion' no existe en el archivo. No se pueden calcular las fechas de vencimiento.")
        return df_local
    
    # Aplicar la función 'calcular_fecha_vencimiento' usando 'Creacion' y 'FechaRenovada'
    df_local["Vencimiento"] = df_local.apply(
        lambda row: calcular_fecha_vencimiento(row["Creacion"], row.get("FechaRenovada", None)),
        axis=1
    )
    
    print("Fechas de vencimiento calculadas y actualizadas para todos los registros.")
    return df_local

def actualizar_status_licencia(df_local):
    """
    Actualiza la columna 'StatusLicencia' basándose en la diferencia entre 
    la fecha actual y la fecha de vencimiento.
    
    Args:
        df_local (pd.DataFrame): DataFrame a actualizar.
    
    Returns:
        pd.DataFrame: DataFrame con la columna 'StatusLicencia' actualizada.
    """
    print("Actualizando la columna 'StatusLicencia'...")
    hoy = datetime.now()
    
    # Convertir la columna 'Vencimiento' a datetime
    df_local["Vencimiento"] = pd.to_datetime(df_local["Vencimiento"], errors="coerce")
    
    # Calcular la diferencia en días entre la fecha actual y la fecha de vencimiento
    df_local["DiasDiferencia"] = (df_local["Vencimiento"] - hoy).dt.days
    
    # Crear la columna 'StatusLicencia' basada en las reglas establecidas
    df_local["StatusLicencia"] = None  # Inicializamos sin valores predeterminados
    df_local.loc[df_local["DiasDiferencia"] > 45, "StatusLicencia"] = "VIGENTE"
    df_local.loc[
        (df_local["DiasDiferencia"] <= 45) & (df_local["DiasDiferencia"] > 30), "StatusLicencia"
    ] = "PRÓXIMO VENCIMIENTO"
    df_local.loc[
        (df_local["DiasDiferencia"] <= 30) & (df_local["DiasDiferencia"] > 0), "StatusLicencia"
    ] = "POR VENCER"
    df_local.loc[df_local["DiasDiferencia"] <= 0, "StatusLicencia"] = "VENCIDA"
    
    # Eliminar la columna auxiliar 'DiasDiferencia' después de usarla
    df_local.drop(columns=["DiasDiferencia"], inplace=True)
    
    print("La columna 'StatusLicencia' ha sido actualizada.")
    return df_local

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

def obtener_registros_activos(df_local):
    """
    Filtra los registros con 'Status' igual a 'Activo' y 
    extrae las columnas 'Serial', 'Empresa', 'Vencimiento' y 'StatusLicencia'.
    Args:
        df_local (pd.DataFrame): DataFrame a filtrar.
    
    Returns:
        pd.DataFrame: DataFrame con los registros activos y las columnas seleccionadas.
    """
    print("Filtrando registros con 'Status' igual a 'Activo'...")
    if "Status" in df_local.columns:
        # Filtrar los registros con 'Status' igual a 'Activo'
        registros_activos = df_local[df_local["Status"] == "Activo"]

        # Seleccionar solo las columnas 'Serial', 'Empresa', 'Vencimiento', 'StatusLicencia'
        columnas_seleccionadas = ["Serial", "Empresa", "Vencimiento", "StatusLicencia"]
        registros_activos = registros_activos[columnas_seleccionadas]
        print(f"Se encontraron {len(registros_activos)} registros activos.")
        return registros_activos
    else:
        print("La columna 'Status' no existe en el archivo.")
        return pd.DataFrame()  # Retorna un DataFrame vacío si no hay datos válidos

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
            # Limpiar la columna 'Contacto'
        df = limpiar_contactos(df)
        # Corregir las direcciones de correo electrónico
        df = corregir_correos(df)

        # Ajustar el formato de la columna 'Empresa'
        df = ajustar_formato_empresa(df)

        # Actualizar las columnas 'FechaActual' y 'HoraActual'
        df = actualizar_fecha_y_hora_actual(df)
        # Calcular fechas de vencimiento para todos los registros
        df["Vencimiento"] = df.apply(
            lambda row: calcular_fecha_vencimiento(row["Creacion"], row.get("FechaRenovada", None)),
            axis=1
        )
        print("Fechas de vencimiento calculadas y actualizadas para todos los registros.")

        # Calcular fechas de vencimiento para todos los registros
        df = calcular_vencimiento(df)

        # Actualizar el estado de la licencia ('StatusLicencia') basado en las fechas
        df = actualizar_status_licencia(df)

        # Depuración: Verificar si df es None después de actualizar las columnas
        if df is None:
            print("El DataFrame es None después de actualizar las columnas. Saliendo del programa.")
            exit()
            # Obtener registros activos
        registros_activos = obtener_registros_activos(df)

        # Mostrar los registros activos
        print("\nRegistros activos:")
        print(registros_activos)


        # Mostrar las primeras filas del archivo
        # print("\nPrimeras filas del archivo:")
        # print(df.head())

        # Mostrar los datos actualizados
        # print("\nDatos actualizados:")
        # print(df[["#", "Empresa", "Creacion", "Vencimiento", "email"]])

        # Guardar los cambios en el archivo Excel
        try:
            df.to_excel(ARCHIVO_EXCEL, index=False)
            print(f"El archivo '{ARCHIVO_EXCEL}' ha sido actualizado correctamente.")
        except PermissionError as e:  # Captura errores de permisos
            print(
                f"No se puede guardar el archivo. Asegúrate de que"
                f"no esté abierto en otro programa: {e}"
                )
        except FileNotFoundError as e:  # Captura errores de archivo no encontrado
            print(
                f"Archivo no encontrado: {e}"
                )
