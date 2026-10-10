"""
Manda el reporte diario de PRUEBA a un solo destinatario, con los datos del
ultimo dia ya procesado (data/acumulado.json -> last_date).

No modifica NADA del estado: arma una copia temporal del acumulado "como estaba
antes de ese dia", genera el mail sobre la copia y lo envia. El workflow que lo
corre no commitea. No usa MAIL_TO: el destinatario llega por argumento.

Uso:  python prueba_diario.py destinatario@dominio.com
"""
import json
import shutil
import sys
import tempfile
from pathlib import Path

from report import build_report, parse_excel_full, BRANCH_ORDER
from send_mail import send

BASE = Path(__file__).parent


def main():
    if len(sys.argv) < 2 or "@" not in sys.argv[1]:
        raise SystemExit("Uso: python prueba_diario.py destinatario@dominio.com")
    destinatario = sys.argv[1].strip()

    estado = json.loads((BASE / "data" / "acumulado.json").read_text(encoding="utf-8"))
    _, fecha, venta_dia, _ = parse_excel_full(BASE / "Ventas_ayer.xlsx")
    if estado.get("last_date") != fecha.isoformat():
        raise SystemExit(f"Ventas_ayer.xlsx es del {fecha} pero el ultimo dia procesado es "
                         f"{estado.get('last_date')}. No mando la prueba para no mostrar numeros cruzados.")

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        (tmp / "data").mkdir()
        shutil.copy(BASE / "data" / f"historico_{fecha.year}.json", tmp / "data")
        # Acumulado "antes de ese dia": el guardado menos la venta del dia.
        previo = dict(estado)
        previo["last_date"] = None
        previo["acumulado"] = {b: round(estado["acumulado"][b] - venta_dia[b], 2) for b in BRANCH_ORDER}
        (tmp / "data" / "acumulado.json").write_text(json.dumps(previo), encoding="utf-8")

        html, _, fecha, totals, _ = build_report(BASE / "Ventas_ayer.xlsx",
                                                 BASE / "Acumulado_interanual.xlsx",
                                                 tmp / "data" / "acumulado.json")

    if abs(totals["a26"] - sum(estado["acumulado"].values())) > 0.01:
        raise SystemExit("El acumulado reconstruido no coincide con el guardado. No se manda.")

    asunto = f"[PRUEBA] Reporte Ventas {fecha.strftime('%d/%m/%Y')} - Lucciano's USA"
    send(asunto, html, to=destinatario)


if __name__ == "__main__":
    main()
