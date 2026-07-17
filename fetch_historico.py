"""
fetch_historico.py
------------------
ONE-SHOT: barre la casilla de Gmail hacia atras, levanta los adjuntos DIARIOS de
TouchBistro que sigan ahi, y llena data/historico_<anio>.json.

Para que sirve:
  - El historial arranca vacio el dia que se implementa. Este script lo siembra
    con lo que todavia este en el mailbox, asi se puede probar el reporte semanal
    con datos reales sin esperar 7 dias.
  - De paso, si llega al 1ro del mes, la conciliacion (historial vs acumulado del
    diario) se prende enseguida en vez de esperar al mes que viene.

Corre a mano desde Actions y despues no se toca mas.

Usa el MISMO criterio que fetch_touchbistro.py para distinguir el diario del
mensual (fecha inicio == fecha fin en el nombre del adjunto) y el MISMO parser
que report.py para consolidar (parse_excel -> VENUE_MAP). Cero logica nueva:
si el diario lee bien un Excel, este tambien.

IDEMPOTENTE: la fecha es la clave del diccionario. Correrlo dos veces no duplica.
NO pisa dias que ya existan salvo que se le pase --pisar.
"""
import argparse
import email
import imaplib
import json
import os
import re
import sys
import tempfile
from datetime import date, datetime, timedelta
from pathlib import Path

from report import parse_excel, BRANCH_ORDER

BASE = Path(__file__).parent
SENDER = os.environ.get("TB_SENDER", "no-reply@touchbistro.com")
FECHAS_RE = re.compile(r"(\d{4}-\d{2}-\d{2})")


def es_diario(filename):
    """True si el adjunto cubre un solo dia. Mismo criterio que fetch_touchbistro.py."""
    fechas = FECHAS_RE.findall(filename or "")
    return len(fechas) >= 2 and fechas[0] == fechas[1]


def cargar(path):
    if not Path(path).exists():
        return {}
    return json.loads(Path(path).read_text(encoding="utf-8"))


def guardar(path, data):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    Path(path).write_text(
        json.dumps({k: data[k] for k in sorted(data)}, indent=2, ensure_ascii=False),
        encoding="utf-8")


def main():
    ap = argparse.ArgumentParser(description="Rescata el historial diario del IMAP.")
    ap.add_argument("--desde", required=True, help="Buscar mails desde esta fecha (YYYY-MM-DD)")
    ap.add_argument("--hasta", help="Ignorar dias posteriores a esta fecha (YYYY-MM-DD)")
    ap.add_argument("--pisar", action="store_true",
                    help="Sobrescribir dias que ya esten en el historial.")
    ap.add_argument("--dry-run", action="store_true",
                    help="Muestra que encontraria pero NO escribe nada.")
    a = ap.parse_args()

    desde = datetime.strptime(a.desde, "%Y-%m-%d").date()
    hasta = datetime.strptime(a.hasta, "%Y-%m-%d").date() if a.hasta else date.today()

    user = os.environ["IMAP_USER"]
    pwd = os.environ["IMAP_APP_PASS"]

    M = imaplib.IMAP4_SSL("imap.gmail.com")
    M.login(user, pwd)
    M.select("INBOX")

    # El mail del dia D llega el dia D+1, asi que busco desde un dia antes por las dudas.
    since = (desde - timedelta(days=1)).strftime("%d-%b-%Y")
    typ, data = M.search(None, f'(FROM "{SENDER}" SINCE {since})')
    ids = data[0].split()
    print(f"[INFO] {len(ids)} mail(s) de {SENDER} desde {since}.")
    if not ids:
        print("[ERROR] No hay mails. Reviso: el remitente es el correcto? La casilla "
              "IMAP_USER es la que recibe los reportes?")
        M.logout()
        return 1

    encontrados = {}   # {fecha: consolidado}
    problemas = []
    tmpdir = Path(tempfile.mkdtemp())

    for n, mid in enumerate(ids, 1):
        typ, msgdata = M.fetch(mid, "(RFC822)")
        msg = email.message_from_bytes(msgdata[0][1])
        for part in msg.walk():
            fn = part.get_filename()
            if not fn or not fn.lower().endswith(".xlsx") or not es_diario(fn):
                continue
            tmp = tmpdir / f"m{n}.xlsx"
            tmp.write_bytes(part.get_payload(decode=True))
            try:
                # La fecha sale del TITULO del Excel, no del nombre del adjunto:
                # el titulo es el dato, el nombre del archivo es una etiqueta.
                f_ini, f_fin, consolidado = parse_excel(tmp)
            except Exception as e:
                problemas.append(f"{fn}: {e}")
                continue
            if f_ini != f_fin:
                continue
            if not (desde <= f_fin <= hasta):
                continue
            encontrados[f_fin] = {b: round(consolidado[b], 2) for b in BRANCH_ORDER}

    M.logout()

    if problemas:
        print(f"\n[AVISO] {len(problemas)} adjunto(s) no se pudieron leer:")
        for p in problemas:
            print(f"  - {p}")
        print("  Causa tipica: una sucursal nueva que falta en VENUE_MAP (report.py).")

    # Reporte de cobertura: que dias del rango pedido estan y cuales no.
    faltantes = []
    d = desde
    while d <= hasta:
        if d not in encontrados:
            faltantes.append(d)
        d += timedelta(days=1)

    print(f"\n[RESULTADO] {len(encontrados)} dia(s) rescatados entre {desde} y {hasta}.")
    if encontrados:
        for f in sorted(encontrados):
            print(f"  OK  {f} = ${sum(encontrados[f].values()):>10,.2f}")
    if faltantes:
        print(f"\n[AVISO] {len(faltantes)} dia(s) del rango NO estan en la casilla:")
        for f in faltantes:
            print(f"  --  {f.isoformat()} ({f.strftime('%a')})")
        print("  Si son dias viejos, probablemente el mail ya no este. Si son recientes, revisar.")

    if a.dry_run:
        print("\n[DRY-RUN] No escribi nada.")
        return 0
    if not encontrados:
        print("\n[ERROR] No rescate ningun dia. No escribo nada.")
        return 1

    # Escribo agrupado por anio (el historial esta separado por anio).
    for anio in sorted({f.year for f in encontrados}):
        path = BASE / "data" / f"historico_{anio}.json"
        hist = cargar(path)
        nuevos = pisados = 0
        for f, v in sorted(encontrados.items()):
            if f.year != anio:
                continue
            k = f.isoformat()
            if k in hist and not a.pisar:
                continue
            pisados += 1 if k in hist else 0
            nuevos += 0 if k in hist else 1
            hist[k] = v
        guardar(path, hist)
        print(f"\n[OK] {path.name}: {nuevos} dia(s) agregado(s), {pisados} pisado(s). "
              f"Total en el archivo: {len(hist)} dias.")

    return 0


if __name__ == "__main__":
    sys.exit(main())
