import pandas as pd
import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Configuración del correo (Vivaldi SMTP)
SMTP_SERVER = "smtp.vivaldi.net"
SMTP_PORT = 465  # Puerto SSL
EMAIL_ADDRESS = "juepae@vivaldi.net"  # Tu correo Vivaldi
EMAIL_PASSWORD = "M1gu3l02.1R3n308.@ndr3s29"     # Tu contraseña o contraseña de aplicación

# Leer el archivo Excel
archivo_excel = "clientes.xlsx"
try:
    df = pd.read_excel(archivo_excel)
except Exception as e:
    print(f"Error al leer el archivo Excel: {e}")
    exit()

# Fecha actual
hoy = datetime.date.today()

# Función para enviar correos en formato HTML
def enviar_correo(destinatario, nombre_cliente, dias_restantes):
    # Asunto del correo
    asunto = f"Notificación de Vencimiento - {dias_restantes} días restantes"

    # Cuerpo del correo en formato HTML
    cuerpo_html = f"""
    <html>
    <head>
        <style>
            body {{
                font-family: Arial, sans-serif;
                background-color: #f4f4f4;
                color: #333;
                padding: 20px;
            }}
            .container {{
                max-width: 600px;
                margin: 0 auto;
                background: #ffffff;
                border-radius: 8px;
                padding: 20px;
                box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            }}
            h1 {{
                color: #d9534f;
                font-size: 24px;
            }}
            p {{
                font-size: 16px;
                line-height: 1.5;
            }}
            .footer {{
                margin-top: 20px;
                font-size: 12px;
                color: #777;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>Notificación de Vencimiento</h1>
            <p>Estimado/a <strong>{nombre_cliente}</strong>,</p>
            <p>Este es un recordatorio de que su anualidad está próxima a vencer en <strong>{dias_restantes} días</strong>.</p>
            <p>Por favor, asegúrese de renovar a tiempo para evitar interrupciones en el servicio.</p>
            <p>Atentamente,<br>Tu Equipo de Soporte</p>
            <div class="footer">
                Este es un mensaje automático. Por favor, no responda a este correo.
            </div>
        </div>
    </body>
    </html>
    """

    # Crear el mensaje
    mensaje = MIMEMultipart()
    mensaje["From"] = EMAIL_ADDRESS
    mensaje["To"] = destinatario
    mensaje["Subject"] = asunto

    # Adjuntar el cuerpo del correo en formato HTML
    mensaje.attach(MIMEText(cuerpo_html, "html"))

    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:  # Usamos SMTP_SSL para SSL
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.sendmail(EMAIL_ADDRESS, destinatario, mensaje.as_string())
        print(f"Notificación enviada a {nombre_cliente} ({destinatario}) para {dias_restantes} días.")
    except Exception as e:
        print(f"Error al enviar el correo a {destinatario}: {e}")

# Procesar cada cliente
for _, row in df.iterrows():
    try:
        nombre_cliente = row["Cliente"]
        correo_cliente = row["Email"]
        fecha_vencimiento = row["Vencido"]

        # Calcular días restantes
        dias_restantes = (fecha_vencimiento - hoy).days

        # Verificar si corresponde enviar una notificación
        if dias_restantes in [30, 15, 10, 5, 4, 3, 2, 1]:
            enviar_correo(correo_cliente, nombre_cliente, dias_restantes)
    except Exception as e:
        print(f"Error al procesar el cliente {nombre_cliente}: {e}")

print("Proceso completado.")