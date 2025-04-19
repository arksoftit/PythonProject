"""Script para probar el envío de correos electrónicos usando SMTP."""

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Configuración del servidor SMTP (Vivaldi)
SMTP_SERVER = "smtp.vivaldi.net"  # Servidor SMTP de Vivaldi
SMTP_PORT = 465  # Puerto SSL (seguro)
EMAIL_ADDRESS = "juepae@vivaldi.net"  # Tu correo Vivaldi
EMAIL_PASSWORD = "M1gu3l02.1R3n308.@ndr3s29"  # Contraseña o contraseña de aplicación

# Dirección del destinatario
RECIPIENT_EMAIL = "a2backup.arksoft@gmail.com"

try:
    # Crear el mensaje
    SUBJECT = "Prueba de Script Python"
    BODY = (
        "El Presente correo es una prueba de mi Script de Python para automatizar los avisos "
        "de vencimientos de clientes, confirme su recepción al correo arksoft.hybrid@outlook.com"
    )

    # Crear el objeto MIME
    message = MIMEMultipart()
    message["From"] = EMAIL_ADDRESS
    message["To"] = RECIPIENT_EMAIL
    message["Subject"] = SUBJECT
    message.attach(MIMEText(BODY, "plain"))

    # Conectar al servidor SMTP (usando SSL)
    print("Conectando al servidor SMTP...")
    server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)  # Usamos SMTP_SSL para el puerto 465
    server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)  # Iniciar sesión en el servidor
    print("Conexión exitosa.")

    # Enviar el correo
    print("Enviando correo...")
    server.sendmail(EMAIL_ADDRESS, RECIPIENT_EMAIL, message.as_string())
    print(f"Correo enviado exitosamente a {RECIPIENT_EMAIL}.")

except smtplib.SMTPException as e:
    print(f"Error al enviar el correo: {e}")

finally:
    # Cerrar la conexión con el servidor
    server.quit()
    print("Conexión cerrada.")