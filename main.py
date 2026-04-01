import pandas as pd
import smtplib
import os
from dotenv import load_dotenv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


# Genera el cuerpo HTML del correo personalizado con el nombre del asociado.
def crear_html(nombre):
    html = f"""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <title>Asamblea General</title>
    </head>
    <body style="margin:0; padding:0; background-color:#f2f2f2; font-family:Arial, sans-serif;">
        <table width="100%" cellpadding="0" cellspacing="0" bgcolor="#f2f2f2">
            <tr>
                <td align="center">
                    <!-- Contenedor -->
                    <table width="450" cellpadding="0" cellspacing="0" bgcolor="#ffffff" style="margin:20px 0;">
                        <!-- Título -->
                        <tr>
                            <td align="center" style="padding:15px; font-weight:bold; color:#333;">
                                Asamblea general
                            </td>
                        </tr>
                        <!-- Banner -->
                        <tr>
                            <td align="center" 
                                style="background:linear-gradient(to right,#1f4e79 0%, #1f4e79 40%, #43c038 100%); padding:50px 20px; color:white;">
                                <h1 style="margin:0;">¡Somos Cooperativa Cóban!</h1>
                                <p style="margin:10px 0 0 0;">Nos alegra tenerte con nosotros</p>
                            </td>
                        </tr>
                        <!-- Contenido -->
                        <tr>
                            <td style="padding:30px; color:#555; font-size:15px;">
                                <h3><strong>{nombre}:</strong><br>
                                <p style="text-align:justify;">
                                    Nos complace informarle que, por haber participado en la votación de la Asamblea General, se 
                                    le ha otorgado un obsequio especial como muestra de nuestro agradecimiento por su compromiso 
                                    y participación activa en nuestra cooperativa.
                                </p>
                                <p style="text-align:justify;">
                                    Complete el siguiente formulario antes del 8 de abril para reclamar su regalo:
                                </p>    
                                <center>
                                <a href="#"
                                   style="background-color:#2e6ddf;
                                          color:#ffffff;
                                          padding:14px 28px;
                                          text-decoration:none;
                                          border-radius:6px;
                                          font-weight:bold;
                                          display:inline-block;">
                                   Enlace para llenar el formulario
                                </a></center>
                                <p style="text-align:justify;">
                                    ¡Gracias por ser parte de nuestra comunidad y por su valiosa participación en la asamblea general!
                                </p>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" style="padding:20px; font-size:12px; color:#999;">
                                © 2026 Cooperativa Cobán. Todos los derechos reservados.
                            </td>
                        </tr>    
                    </table>
                </td>
            </tr>
        </table>
    </body>
    </html>
    """
    return html

def main():
    # Cargar variables de entorno
    load_dotenv()
    # Credenciales del remitente tomadas desde variables de entorno.
    correo = os.getenv('EMAIL_USER')
    contraseña = os.getenv('EMAIL_PASS')
    # Valida que existan las credenciales antes de continuar.
    if not correo or not contraseña:
        raise ValueError('Faltan EMAIL_USER o EMAIL_PASS en el archivo .env')

    # Leer Excel
    df = pd.read_excel('votos_linea.xlsx')
    # Lista de correos
    correos = []

    for _, row in df.iterrows():
        email = row['Correo Electronico']

        if pd.isna(email):
            continue

        correos.append(str(email).strip())

    # Validación: evitar enviar si no hay correos
    if not correos:
        print("No hay correos válidos para enviar")
        return

    # Configurar servidor SMTP
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(correo, contraseña)

    try:
        msg = MIMEMultipart()
        msg['From'] = correo
        msg['To'] = correo 
        msg['Bcc'] = ", ".join(correos)
        msg['Subject'] = 'Gracias por votar'

        html_content = crear_html("Estimado asociado")
        msg.attach(MIMEText(html_content, 'html'))

        server.send_message(msg)

        print(f"Correo enviado a {len(correos)} destinatarios")

    finally:
        server.quit()


if __name__ == '__main__':
    main()