import pandas as pd
import smtplib
import os
from dotenv import load_dotenv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def crear_html(nombre):
    html = f"""
    <!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Asamblea General</title>
</head>
<body style="margin:0; padding:0; font-family: Arial, sans-serif; background-color:#f2f2f2;">

    <table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr>
            <td align="center">
                
                <!-- Contenedor -->
                <table width="600" cellpadding="0" cellspacing="0" style="background:#ffffff; margin-top:20px; border-radius:8px; overflow:hidden;">
                    
                    <!-- Título -->
                    <tr>
                        <td align="center" style="padding:15px; font-weight:bold; color:#333;">
                            Asamblea general
                        </td>
                    </tr>

                    <!-- Banner -->
                    <tr>
                        <td style="background: linear-gradient(90deg, #2c5d8a, #3bb143); color:white; text-align:center; padding:40px 20px;">
                            <h1 style="margin:0;">¡Somos Cooperativa Cobán!</h1>
                            <p style="margin:10px 0 0 0;">Nos alegra tenerte con nosotros</p>
                        </td>
                    </tr>

                    <!-- Contenido -->
                    <tr>
                        <td style="padding:30px; color:#444; line-height:1.6;">
                            
                            <p><strong>Estimado asociado: {nombre}</strong></p>

                            <p>
                                Nos complace informarle que por aver participado en la votación de la asamblea general, se 
                                le ha otorgado un regalo especial como muestra de nuestro agradecimiento por su compromiso 
                                y participación activa en nuestra cooperativa.
                            </p>

                            <p>
                                Llene el siguiente formulario hastal el 8 de abril para reclamar su regalo:
                            </p>

                            <a href="#" target="_blank" style="background-color: #2c5d8a; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">Llenar formulario</a>

                            <p>
                                ¡Gracias por ser parte de nuestra comunidad y por su valiosa participación en la asamblea general!
                            </p>

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
    load_dotenv()

    correo = os.getenv('EMAIL_USER')
    contraseña = os.getenv('EMAIL_PASS')

    if not correo or not contraseña:
        raise ValueError('Faltan EMAIL_USER o EMAIL_PASS en el archivo .env')

    df = pd.read_excel('votos_linea.xlsx')

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(correo, contraseña)

    try:
        for _, row in df.iterrows():
            nombre = row['Nombre']
            email = row['Correo Electronico']

            msg = MIMEMultipart()
            msg['From'] = correo
            msg['To'] = email
            msg['Subject'] = 'Gracias por votar'

            html_content = crear_html(nombre)
            msg.attach(MIMEText(html_content, 'html'))

            server.send_message(msg)
    finally:
        server.quit()

    print('Correos enviados')


if __name__ == '__main__':
    main()