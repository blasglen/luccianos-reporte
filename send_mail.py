"""
Envia el reporte HTML por Gmail via SMTP usando un App Password.
Incrusta imagenes inline (Content-ID) referenciadas en el HTML como cid:<nombre>:
  - cid:logo        -> Logo.png
  - cid:comparativo -> charts/comparativo.png
  - cid:progreso    -> charts/progreso.png

Variables de entorno (inyectadas por GitHub Actions desde Secrets):
  GMAIL_USER      -> casilla emisora (la cuenta gmail que envia)
  GMAIL_APP_PASS  -> App Password de 16 caracteres (NO la clave normal)
  MAIL_TO         -> destinatario(s). Si hay varios, separados por coma.
                     El PRIMERO va en 'Para' (To); el resto en copia (CC).
"""
import os
import smtplib
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from pathlib import Path

# (cid, ruta)
INLINE_IMAGES = [
    ("logo", "Logo.png"),
    ("comparativo", "charts/comparativo.png"),
    ("progreso", "charts/progreso.png"),
]


def _attach_inline(root, cid, path):
    p = Path(path)
    if not p.exists():
        print(f"[AVISO] No encontre {path}; el mail sale sin esa imagen (cid:{cid}).")
        return
    img = MIMEImage(p.read_bytes())
    img.add_header("Content-ID", f"<{cid}>")
    img.add_header("Content-Disposition", "inline", filename=p.name)
    root.attach(img)


def send(subject, html, to=None):
    user = os.environ["GMAIL_USER"]
    pwd = os.environ["GMAIL_APP_PASS"]
    raw = to or os.environ["MAIL_TO"]

    destinatarios = [d.strip() for d in raw.split(",") if d.strip()]
    if not destinatarios:
        raise ValueError("MAIL_TO no tiene ninguna direccion valida.")

    to_addr = destinatarios[0]
    cc_addrs = destinatarios[1:]
    all_rcpts = destinatarios

    root = MIMEMultipart("related")
    root["Subject"] = subject
    root["From"] = f"Lucciano's USA <{user}>"
    root["To"] = to_addr
    if cc_addrs:
        root["Cc"] = ", ".join(cc_addrs)

    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText("Tu cliente no soporta HTML. Abri el reporte en un cliente compatible.", "plain", "utf-8"))
    alt.attach(MIMEText(html, "html", "utf-8"))
    root.attach(alt)

    for cid, path in INLINE_IMAGES:
        _attach_inline(root, cid, path)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
        s.login(user, pwd)
        s.sendmail(user, all_rcpts, root.as_string())
    print(f"Mail enviado. Para: {to_addr} | CC: {', '.join(cc_addrs) if cc_addrs else '(ninguno)'}")


if __name__ == "__main__":
    subject = sys.argv[1] if len(sys.argv) > 1 else "Reporte de Ventas"
    path = sys.argv[2] if len(sys.argv) > 2 else "preview.html"
    with open(path, encoding="utf-8") as f:
        send(subject, f.read())
