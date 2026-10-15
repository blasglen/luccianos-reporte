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


def es_mensual(filename, anio, mes):
    """True si el adjunto cubre el mes completo anio-mes (del 1 al ultimo dia)."""
    import calendar
    fechas = FECHAS_RE.findall(filename or "")
    ult = calendar.monthrange(anio, mes)[1]
    return (len(fechas) >= 2 and fechas[0] == f"{anio:04d}-{mes:02d}-01"
            and fechas[1] == f"{anio:04d}-{mes:02d}-{ult:02d}")


def main_mensual():
    """Baja el reporte MENSUAL del mes anterior (mismo 'Sales Summary', rango de
    un mes completo) a Ventas_mes.xlsx. Lo usa el mail de sales tax del dia 1."""
    hoy = datetime.utcnow().date()
    anio, mes = (hoy.year, hoy.month - 1) if hoy.month > 1 else (hoy.year - 1, 12)
    M = imaplib.IMAP4_SSL("imap.gmail.com")
    M.login(os.environ["IMAP_USER"], os.environ["IMAP_APP_PASS"])
    M.select("INBOX")
    since = (datetime.utcnow() - timedelta(days=5)).strftime("%d-%b-%Y")
    typ, data = M.search(None, f'(FROM "{SENDER}" SINCE {since})')
    for mid in reversed(data[0].split()):
        typ, msgdata = M.fetch(mid, "(RFC822)")
        msg = email.message_from_bytes(msgdata[0][1])
        for part in msg.walk():
            fn = part.get_filename()
            if fn and fn.lower().endswith(".xlsx") and es_mensual(fn, anio, mes):
                destino = Path(__file__).parent / "Ventas_mes.xlsx"
                destino.write_bytes(part.get_payload(decode=True))
                print(f"[OK] Guardado {destino.name} desde adjunto {fn}")
                M.logout()
                return 0
    M.logout()
    print(f"[ERROR] No llego el reporte mensual de {anio}-{mes:02d} de TouchBistro "
          f"(adjunto del 1 al ultimo dia del mes).")
    return 1


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
    # Sin argumentos: el diario de siempre. Con --mensual: el del mes anterior.
    sys.exit(main_mensual() if "--mensual" in sys.argv else main())
