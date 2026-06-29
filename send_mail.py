"""
Envia el reporte HTML por Gmail via SMTP usando un App Password.
Lee de variables de entorno (inyectadas por GitHub Actions desde Secrets):
  GMAIL_USER      -> casilla emisora (ej: administracion@luccianos.com.ar o la cuenta gmail)
  GMAIL_APP_PASS  -> App Password de 16 caracteres (NO la clave normal)
  MAIL_TO         -> destinatario fijo
"""
import os
import smtplib
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def send(subject, html, to=None):
    user = os.environ["GMAIL_USER"]
    pwd = os.environ["GMAIL_APP_PASS"]
    to = to or os.environ["MAIL_TO"]

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = f"Lucciano's Reportes <{user}>"
    msg["To"] = to
    msg.attach(MIMEText("Tu cliente no soporta HTML. Abrí el reporte en un cliente compatible.", "plain", "utf-8"))
    msg.attach(MIMEText(html, "html", "utf-8"))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
        s.login(user, pwd)
        s.sendmail(user, [to], msg.as_string())
    print(f"Mail enviado a {to}")


if __name__ == "__main__":
    # Uso: send_mail.py "Asunto" archivo.html
    subject = sys.argv[1] if len(sys.argv) > 1 else "Reporte de Ventas"
    path = sys.argv[2] if len(sys.argv) > 2 else "preview.html"
    with open(path, encoding="utf-8") as f:
        send(subject, f.read())
