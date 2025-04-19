"""
Script para automatizar el envío de correos electrónicos usando SMTP.
Este script carga credenciales desde un archivo .env y envía correos de prueba.
"""

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
from dotenv import load_dotenv
from datetime import datetime

# Cargar las variables de entorno desde el archivo .env
load_dotenv()

# Configuración segura usando variables de entorno
SMTP_SERVER = "smtp.vivaldi.net"  # Servidor SMTP de Vivaldi
SMTP_PORT = 465  # Puerto SSL (seguro)
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")  # Tu correo Vivaldi (desde .env)
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")  # Contraseña o contraseña de aplicación (desde .env)

def enviar_correo(destinatario, asunto, cuerpo):
    """
    Función para enviar un correo electrónico a través de SMTP.
    
    Args:
        destinatario (str): Dirección de correo del destinatario.
        asunto (str): Asunto del correo.
        cuerpo (str): Cuerpo del correo.
    
    Returns:
        bool: True si el correo fue enviado exitosamente, False en caso contrario.
    """
    try:
        # Crear el objeto MIME
        message = MIMEMultipart()
        message["From"] = EMAIL_ADDRESS
        message["To"] = destinatario
        message["Subject"] = asunto
        message.attach(MIMEText(cuerpo, "plain"))

        # Conectar al servidor SMTP (usando SSL)
        print("Conectando al servidor SMTP...")
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)  # Iniciar sesión en el servidor
            print("Conexión exitosa.")

            # Enviar el correo
            print(f"Enviando correo a {destinatario}...")
            server.sendmail(EMAIL_ADDRESS, destinatario, message.as_string())
            print(f"Correo enviado exitosamente a {destinatario}.")
        
        return True

    except (smtplib.SMTPException, smtplib.SMTPAuthenticationError, smtplib.SMTPConnectError) as e:
        print(f"Error SMTP al enviar el correo: {e}")
        return False


def obtener_fecha_formateada():
    # Obtener la fecha actual
    fecha_actual = datetime.now()

    # Formatear la fecha como "Caracas; Día de Mes de Año"
    meses = [
        "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
    ]
    dia = fecha_actual.day
    mes = meses[fecha_actual.month - 1]  # Los meses están indexados desde 0
    año = fecha_actual.year

    return f"Caracas; {dia} de {mes} de {año}."


if __name__ == "__main__":
    # Ejemplo de uso
    RECIPIENT_EMAIL = "a2backup.arksoft@gmail.com"
    SUBJECT = "Prueba de Script Python"

    # Obtener la fecha formateada
    fecha_formateada = obtener_fecha_formateada()

    BODY = (
        f"{fecha_formateada}\n\n"
        "El presente correo es una prueba de mi Script de Python para automatizar "
        "los avisos de vencimientos de clientes. Confirme su recepción al correo "
        "arksoft.hybrid@outlook.com."
    )

    # Verificar si las variables de entorno están configuradas
    if not EMAIL_ADDRESS or not EMAIL_PASSWORD:
        print("Error: Las variables de entorno EMAIL_ADDRESS y EMAIL_PASSWORD deben estar configuradas.")
    else:
        # Enviar correo
        ENVIADO = enviar_correo(RECIPIENT_EMAIL, SUBJECT, BODY)
        if ENVIADO:
            print("Proceso completado exitosamente.")
        else:
            print("El proceso terminó con errores.")