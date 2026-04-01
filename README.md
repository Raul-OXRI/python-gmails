# Envío de correos masivos con Python y Gmail

Script en Python para enviar un correo HTML a múltiples destinatarios usando Gmail SMTP.

## Qué hace este proyecto

- Lee una lista de correos desde el archivo `Reprote votos virtuales.csv`.
- Toma credenciales de Gmail desde variables de entorno (`.env`).
- Genera un correo en formato HTML.
- Envía un único correo con destinatarios en `Bcc` (copia oculta).

## Estructura del proyecto

- `main.py`: lógica principal de lectura de Excel y envío de correos.
- `Reprote votos virtuales.csv`: archivo con los correos destino.
- `.env`: credenciales del remitente (no se comparte en git).

## Requisitos

- Python 3.9 o superior.
- Cuenta de Gmail con App Password activa.

Dependencias de Python:

- pandas
- openpyxl
- python-dotenv

## Instalación

1. Crear y activar entorno virtual (opcional, recomendado):

```powershell
python -m venv venv
venv\Scripts\activate
```

2. Instalar dependencias:

```powershell
pip install pandas openpyxl python-dotenv
```
o
```
pip install -r requirements.txt
```

## Configuración

1. Crear el archivo `.env` en la raíz del proyecto.
2. Agregar estas variables:

```env
EMAIL_USER=tu_correo@gmail.com
EMAIL_PASS=tu_app_password
```

Notas importantes:

- No uses tu contraseña normal de Gmail.
- Usa una App Password de Google (con verificación en dos pasos habilitada).

## Formato del csv

El script espera un archivo llamado `Reprote votos virtuales.csv` con una columna exacta:

- `Correo Electronico`

Si una fila no tiene correo, se omite automáticamente.

## Ejecución

```powershell
python main.py
```

Salida esperada:

- `Correo enviado a N destinatarios`, si todo sale bien.
- `No hay correos válidos para enviar`, si no se encontraron destinatarios.

## Cómo personalizar el correo

Puedes modificar la plantilla HTML en la función `crear_html(nombre)` dentro de `main.py`.

Puntos comunes para personalizar:

- Asunto del correo (`msg['Subject']`).
- Texto del mensaje HTML.

## Errores frecuentes

- **Faltan EMAIL_USER o EMAIL_PASS en el archivo .env**: revisa que el archivo `.env` exista y tenga ambas variables.

- **Error de autenticación SMTP**: verifica que `EMAIL_PASS` sea una App Password válida.

- **Error al leer Excel**: confirma que `votos_linea.xlsx` exista en la raíz del proyecto y tenga la columna `Correo Electronico`.

## Seguridad

- **Nunca compartas el archivo `.env`**: contiene credenciales sensibles. Agrégalo a `.gitignore`.
- **Usa App Passwords**: no pongas tu contraseña principal de Google en el código.
- **Verifica destinatarios**: revisa tu Excel antes de ejecutar para evitar envíos accidentales.
- **Respeta límites de Gmail**: Gmail permite enviar hasta ~500 correos/día desde scripts SMTP.
