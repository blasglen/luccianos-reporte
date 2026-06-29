"""
Envia el reporte HTML por Gmail via SMTP usando un App Password.
El logo se incrusta como adjunto inline (Content-ID: logo), referenciado en el HTML como cid:logo.

Variables de entorno (inyectadas por GitHub Actions desde Secrets):
  GMAIL_USER      -> casilla emisora (la cuenta gmail que envia)
  GMAIL_APP_PASS  -> App Password de 16 caracteres (NO la clave normal)
  MAIL_TO         -> destinatario fijo
"""
import os
import smtplib
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from pathlib import Path


def send(subject, html, to=None, logo_path="Logo.png"):
    user = os.environ["GMAIL_USER"]
    pwd = os.environ["GMAIL_APP_PASS"]
    to = to or os.environ["MAIL_TO"]

    # 'related' permite incrustar imagenes referenciadas por cid: dentro del HTML.
    root = MIMEMultipart("related")
    root["Subject"] = subject
    root["From"] = f"Lucciano's Reportes <{user}>"
    root["To"] = to

    # Parte alternativa: texto plano + HTML
    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText("Tu cliente no soporta HTML. Abri el reporte en un cliente compatible.", "plain", "utf-8"))
    alt.attach(MIMEText(html, "html", "utf-8"))
    root.attach(alt)

    # Logo inline (cid:logo)
    p = Path(logo_path)
    if p.exists():
        img = MIMEImage(p.read_bytes())
        img.add_header("Content-ID", "<logo>")
        img.add_header("Content-Disposition", "inline", filename="Logo.png")
        root.attach(img)
    else:
        print(f"[AVISO] No encontre {logo_path}; el mail sale sin logo.")

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
        s.login(user, pwd)
        s.sendmail(user, [to], root.as_string())
    print(f"Mail enviado a {to}")


if __name__ == "__main__":
    subject = sys.argv[1] if len(sys.argv) > 1 else "Reporte de Ventas"
    path = sys.argv[2] if len(sys.argv) > 2 else "preview.html"
    with open(path, encoding="utf-8") as f:
        send(subject, f.read())
