"""
sales_tax.py — Mail mensual para preparar el SALES TAX (Florida, DR-15) de las
sucursales propias. Va solo a admin@luccianos.us (Secret MAIL_TO_SALES_TAX).

De donde salen los numeros:
  El 1ro de cada mes TouchBistro manda el MISMO reporte "Sales Summary" del diario,
  pero con el mes completo (fecha inicio = dia 1, fecha fin = ultimo dia). Ese
  cuadro NO trae el impuesto cobrado: el impuesto se CALCULA aca como
      base imponible = Net Sales del mes (sin impuesto, despues de descuentos)
      impuesto       = base x (6% estatal + surtax del condado)

Control: la base de cada local se compara contra la suma del historial diario
(data/historico_<anio>.json). Si no cierran al centavo, el mail sale igual pero
con un aviso arriba, para revisar antes de declarar.

Uso:
    python sales_tax.py Ventas_mes.xlsx         -> arma preview_sales_tax.html
    python sales_tax.py --desde-historial 2026-09   (borrador sin el Excel mensual)
"""
import json
import os
import sys
from datetime import date
from pathlib import Path

import openpyxl

from report import VENUE_MAP, MESES_ES

BASE = Path(__file__).parent
PREVIEW = BASE / "preview_sales_tax.html"

# Sucursales propias que declaran sales tax, con su condado y surtax.
# Fuente de tasas: Florida DOR, DR-15DSS (Orange 0.5%, Broward 1%).
# Si cambian las propias o una tasa, se toca SOLO esto.
ESTATAL = 0.06
SUCURSALES = {
    "Florida Mall": {"condado": "Orange",  "surtax": 0.005},
    "Vineland":     {"condado": "Orange",  "surtax": 0.005},
    "Weston":       {"condado": "Broward", "surtax": 0.010},
}

# Datos de registro de cada local (razon social, ubicacion, Certificate # y
# Business Partner #). NO van en el codigo: el repo es publico y esos numeros
# sirven para entrar al portal del DOR. Se leen del Secret SALES_TAX_DATOS (JSON):
#   {"Florida Mall": {"entidad": "...", "ubicacion": "...", "cert": "...", "bp": "..."}, ...}
LOGIN_DOR = "https://floridarevenue.com/taxes/eservices/Pages/sutlogin.aspx"


def datos_registro():
    raw = os.environ.get("SALES_TAX_DATOS", "").strip()
    if not raw:
        print("[AVISO] Falta el Secret SALES_TAX_DATOS: el mail sale sin certificado ni Business Partner.")
        return {}
    try:
        return json.loads(raw)
    except ValueError:
        raise SystemExit("[ERROR] El Secret SALES_TAX_DATOS no es un JSON valido.")


def _f(v):
    if v is None:
        return 0.0
    return float(str(v).replace(",", "").replace("$", "").strip() or 0)


def leer_mensual(path):
    """Lee el Sales Summary mensual. Devuelve (desde, hasta, {sucursal: datos})."""
    import re
    from datetime import datetime
    ws = openpyxl.load_workbook(path, data_only=True).active
    filas = list(ws.iter_rows(values_only=True))
    m = re.search(r"(\d{4}-\d{2}-\d{2})\s*/\s*(\d{4}-\d{2}-\d{2})", str(filas[0][0] or ""))
    if not m:
        raise SystemExit(f"[ERROR] No pude leer el rango de fechas del titulo de {path}.")
    desde = datetime.strptime(m.group(1), "%Y-%m-%d").date()
    hasta = datetime.strptime(m.group(2), "%Y-%m-%d").date()
    datos = {}
    for r in filas[2:]:
        nombre = str(r[0] or "").strip()
        if not nombre or nombre.upper().startswith("REPORT"):
            continue
        if nombre not in VENUE_MAP:
            raise SystemExit(f"[ERROR] Venue desconocido: {nombre!r}. Agregalo a VENUE_MAP.")
        b = VENUE_MAP[nombre]
        d = datos.setdefault(b, {"bruta": 0.0, "neta": 0.0, "descuentos": 0.0, "anulados": 0.0, "tickets": 0})
        d["bruta"] += _f(r[1]); d["neta"] += _f(r[2]); d["descuentos"] += _f(r[3])
        d["anulados"] += _f(r[4]); d["tickets"] += int(_f(r[5]))
    return desde, hasta, datos


def desde_historial(anio, mes):
    """Borrador sin el Excel mensual: suma el historial diario (solo venta neta y tickets)."""
    import calendar
    hist = json.loads((BASE / "data" / f"historico_{anio}.json").read_text(encoding="utf-8"))
    ult = calendar.monthrange(anio, mes)[1]
    datos = {}
    for d in range(1, ult + 1):
        dia = hist.get(date(anio, mes, d).isoformat())
        if dia is None:
            raise SystemExit(f"[ERROR] Falta el {date(anio, mes, d)} en el historial.")
        for b, v in dia.items():
            x = datos.setdefault(b, {"bruta": None, "neta": 0.0, "descuentos": None, "anulados": None, "tickets": 0})
            x["neta"] += v["venta"]; x["tickets"] += v["tickets"]
    return date(anio, mes, 1), date(anio, mes, ult), datos


def control_historial(desde, hasta, datos):
    """Compara la Net Sales de cada local contra la suma del historial diario."""
    p = BASE / "data" / f"historico_{desde.year}.json"
    if not p.exists():
        return ["No hay historial diario de ese año para controlar."]
    hist = json.loads(p.read_text(encoding="utf-8"))
    avisos = []
    for b in SUCURSALES:
        suma, faltan = 0.0, 0
        for n in range((hasta - desde).days + 1):
            dia = hist.get(date.fromordinal(desde.toordinal() + n).isoformat())
            if dia is None:
                faltan += 1
                continue
            suma += dia[b]["venta"]
        if faltan:
            avisos.append(f"Al historial le faltan {faltan} día(s): no se pudo controlar {b}.")
        elif abs(suma - datos[b]["neta"]) > 0.01:
            avisos.append(f"{b}: el reporte mensual dice ${datos[b]['neta']:,.2f} y la suma de los "
                          f"diarios ${suma:,.2f} (diferencia ${datos[b]['neta'] - suma:,.2f}).")
    return avisos


def calcular(datos):
    reg = datos_registro()
    filas = []
    for b, cfg in SUCURSALES.items():
        if b not in datos:
            raise SystemExit(f"[ERROR] {b} no aparece en el reporte mensual.")
        d = datos[b]
        base = round(d["neta"], 2)
        estatal = round(base * ESTATAL, 2)
        surtax = round(base * cfg["surtax"], 2)
        filas.append({"suc": b, "reg": reg.get(b, {}), **cfg, **d, "base": base, "estatal": estatal, "sur": surtax,
                      "total": round(estatal + surtax, 2), "tasa": ESTATAL + cfg["surtax"]})
    return filas


def render(desde, hasta, filas, avisos):
    from mail_base import header, tarjeta, pie, documento, plata, num, callout
    mes = MESES_ES[desde.month].capitalize()
    tot = {k: round(sum(f[k] for f in filas), 2) for k in ("base", "estatal", "sur", "total")}

    def linea(lbl, val, fuerte=False, mut=False):
        st = "font-weight:800;font-size:16px;" if fuerte else ""
        cl = ' class="mut"' if mut else ""
        return (f'<tr><td{cl} style="padding:6px 0;border-bottom:1px solid #eae7e1;font-size:14px;{st}">{lbl}</td>'
                f'<td{cl} style="padding:6px 0;border-bottom:1px solid #eae7e1;font-size:14px;text-align:right;'
                f'white-space:nowrap;{st}">{val}</td></tr>')

    bloques = ""
    for f in filas:
        extra = ""
        if f["bruta"] is not None:
            extra = (linea("Venta bruta", plata(f["bruta"], 2), mut=True)
                     + linea("Descuentos", plata(-f["descuentos"], 2), mut=True))
        r = f["reg"]
        titulo = f'{r["entidad"]} – {r["ubicacion"]}' if r.get("entidad") else f["suc"]
        ids = ""
        if r.get("cert"):
            ids = (f'<table role="presentation" style="width:100%;margin-top:10px;background:#f5f3ef;border-radius:8px"><tr>'
                   f'<td style="padding:8px 10px"><div class="mut" style="font-size:11px">Certificate #</div>'
                   f'<div class="b" style="font-size:14px;letter-spacing:.3px">{r["cert"]}</div></td>'
                   f'<td style="padding:8px 10px;text-align:right"><div class="mut" style="font-size:11px">Business Partner #</div>'
                   f'<div class="b" style="font-size:14px;letter-spacing:.3px">{r["bp"]}</div></td></tr></table>')
        bloques += f'''
 <div class="sec">
  <table role="presentation" style="width:100%"><tr>
    <td><div class="h">{titulo}</div><div class="mut sm">{f["condado"]} County · tasa {f["tasa"] * 100:.1f}% · {num(f["tickets"])} tickets</div></td>
    <td style="text-align:right;vertical-align:top"><div class="lbl" style="margin:0">A pagar</div><div class="big2" style="font-size:24px">{plata(f["total"], 2)}</div></td>
  </tr></table>
  {ids}
  <table role="presentation" style="width:100%;margin-top:10px">
   {extra}
   {linea("Ventas netas = base imponible", plata(f["base"], 2), fuerte=True)}
   {linea("Impuesto estatal 6%", plata(f["estatal"], 2))}
   {linea(f"Surtax {f['condado']} {f['surtax'] * 100:.1f}%", plata(f["sur"], 2))}
  </table>
 </div>'''

    aviso = "".join(callout(f"⚠ {a}") for a in avisos)
    resumen = f'''
 <div class="sec">
  <div class="lbl">Sales tax {mes.lower()} {desde.year} · {len(filas)} sucursales propias</div>
  <div class="big2">{plata(tot["total"], 2)}</div>
  <div class="msub">Sobre ventas netas por {plata(tot["base"], 2)}: {plata(tot["estatal"], 2)} de impuesto estatal y {plata(tot["sur"], 2)} de surtax de condado. Período {desde.strftime("%d/%m")} al {hasta.strftime("%d/%m/%Y")}.</div>
  {aviso}
  <table role="presentation" style="margin-top:14px"><tr><td style="background:#1a1a1a;border-radius:8px;padding:10px 16px">
    <a href="{LOGIN_DOR}" style="color:#ffffff;font-weight:700;font-size:14px;text-decoration:none">Ingresar al portal de Florida DOR →</a>
  </td></tr></table>
 </div>'''
    nota = '''
 <div class="sec">
  <div class="nota">El impuesto está <b>calculado</b> sobre la venta neta del reporte mensual de TouchBistro (Sales Summary), que no informa el impuesto cobrado. Tasas: 6% estatal + surtax del condado según el formulario DR-15DSS del Florida Department of Revenue. No incluye el descuento por presentación en término (collection allowance).</div>
 </div>'''
    card = tarjeta(header("Sales tax", f"{mes} {desde.year}"), resumen, bloques, nota,
                   pie=pie(f"Lucciano's USA · Sucursales propias · {desde.strftime('%d/%m')} al {hasta.strftime('%d/%m/%Y')}"))
    return documento(card, f"Sales tax {mes} {desde.year}", f"Sales tax {mes.lower()}: {plata(tot['total'], 2)} a pagar")


def main():
    if len(sys.argv) >= 3 and sys.argv[1] == "--desde-historial":
        anio, mes = map(int, sys.argv[2].split("-"))
        desde, hasta, datos = desde_historial(anio, mes)
        avisos = []
    else:
        path = sys.argv[1] if len(sys.argv) > 1 else str(BASE / "Ventas_mes.xlsx")
        desde, hasta, datos = leer_mensual(path)
        if desde.day != 1 or hasta.month != desde.month:
            raise SystemExit(f"[ERROR] El reporte no es de un mes completo: {desde}..{hasta}.")
        avisos = control_historial(desde, hasta, datos)
    filas = calcular(datos)
    PREVIEW.write_text(render(desde, hasta, filas, avisos), encoding="utf-8")
    asunto = f"Sales tax {MESES_ES[desde.month].capitalize()} {desde.year} - Lucciano's USA"
    gh = os.environ.get("GITHUB_OUTPUT")
    if gh:
        with open(gh, "a", encoding="utf-8") as f:
            f.write(f"subject={asunto}\n")
    for f in filas:
        print(f"{f['suc']:13} base ${f['base']:>11,.2f}  tax ${f['total']:>9,.2f}")
    print(f"TOTAL tax ${sum(f['total'] for f in filas):,.2f}" + (f"  | {len(avisos)} aviso(s)" if avisos else ""))
    return 0


if __name__ == "__main__":
    sys.exit(main())
