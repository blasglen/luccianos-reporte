"""
fetch_touchbistro.py
--------------------
Lee la casilla de Gmail por IMAP, busca el ultimo mail de TouchBistro con el
reporte diario, y guarda el adjunto como Ventas_ayer.xlsx.

Usa App Password (mismo esquema que send_mail.py), NO OAuth.
Variables de entorno (Secrets de GitHub Actions):
  IMAP_USER      -> la casilla que RECIBE el reporte de TouchBistro
  IMAP_APP_PASS  -> App Password de 16 caracteres de ESA casilla

Robustez:
  - El mail diario y el mensual llegan con el mismo asunto ("Your report is here").
    Los distingo por el nombre del adjunto: el diario tiene rango de UN dia
    (fecha inicio == fecha fin, ej. ...2026-06-30-2026-06-30...). Ese es el que agarro.
  - Falla RUIDOSA: si no encuentra nada, corta con error para que el job quede en rojo.
"""
import email
import imaplib
import os
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

SENDER = os.environ.get("TB_SENDER", "no-reply@touchbistro.com")
SALIDA = Path(__file__).parent / "Ventas_ayer.xlsx"

FECHAS_RE = re.compile(r"(\d{4}-\d{2}-\d{2})")


def es_diario(filename):
    """True si el adjunto cubre un solo dia (fecha inicio == fecha fin)."""
    fechas = FECHAS_RE.findall(filename or "")
    return len(fechas) >= 2 and fechas[0] == fechas[1]


def main():
    user = os.environ["IMAP_USER"]
    pwd = os.environ["IMAP_APP_PASS"]

    M = imaplib.IMAP4_SSL("imap.gmail.com")
    M.login(user, pwd)
    M.select("INBOX")

    # Ultimos 2 dias, del remitente de TouchBistro.
    since = (datetime.utcnow() - timedelta(days=2)).strftime("%d-%b-%Y")
    typ, data = M.search(None, f'(FROM "{SENDER}" SINCE {since})')
    ids = data[0].split()
    if not ids:
        print(f"[ERROR] No hay mails de {SENDER} en los ultimos 2 dias.")
        return 1

    fallback = None  # primer xlsx que aparezca, por si ninguno es "diario"
    # Recorro de mas nuevo a mas viejo.
    for mid in reversed(ids):
        typ, msgdata = M.fetch(mid, "(RFC822)")
        msg = email.message_from_bytes(msgdata[0][1])
        for part in msg.walk():
            fn = part.get_filename()
            if not fn or not fn.lower().endswith(".xlsx"):
                continue
            payload = part.get_payload(decode=True)
            if fallback is None:
                fallback = payload
            if es_diario(fn):
                SALIDA.write_bytes(payload)
                print(f"[OK] Guardado {SALIDA.name} desde adjunto {fn} ({len(payload)} bytes)")
                M.logout()
                return 0

    if fallback is not None:
        SALIDA.write_bytes(fallback)
        print(f"[AVISO] No encontre un adjunto de un solo dia; use el mas reciente "
              f"({len(fallback)} bytes). Revisa que sea el diario correcto.")
        M.logout()
        return 0

    M.logout()
    print("[ERROR] Habia mails de TouchBistro pero ninguno con adjunto .xlsx.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
