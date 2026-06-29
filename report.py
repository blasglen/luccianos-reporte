"""
Lucciano's - Reporte de Ventas Diario
Genera el mail HTML comparando el mes en curso (2026) contra el mismo periodo del anio anterior (2025).

Flujo:
  1. Lee Ventas_ayer.xlsx       -> venta del dia (2026)
  2. Lee acum_jun26.json        -> acumulado del mes previo (persistido en el repo)
  3. Acum.26 = acum previo + venta del dia  (se vuelve a guardar)
  4. Lee Acumulado_interanual.xlsx -> acumulado 2025 (comparativo)
  5. Variacion = (acum26 - acum25) / acum25
  6. Genera el HTML del mail
"""
import json
import os
import re
import sys
from pathlib import Path
from datetime import datetime, timedelta

import openpyxl

# --- Configuracion de sucursales -------------------------------------------------
# Mapea las filas crudas del Excel (7 venues) a las 6 sucursales consolidadas del mail.
# Las dos Vineland se suman en una sola.
VENUE_MAP = {
    "#001 Florida Mall Orlando FL": "Florida Mall",
    "#004 Weston Town Center FL": "Weston",
    "#005 Vineland Orlando FL": "Vineland",
    "Lucciano's Vineland 2026": "Vineland",
    "#002 American Dream Mall NJ": "American Dream",
    "#003 Sawgrass Mills Mall FL": "Sawgrass",
    "#006 Aventura, FL": "Aventura",
}

# Orden de aparicion en la tabla de detalle
BRANCH_ORDER = ["Florida Mall", "Weston", "Vineland", "American Dream", "Sawgrass", "Aventura"]

# Sucursales "PROPIAS" (seccion destacada en el mail)
PROPIAS = ["Florida Mall", "Weston", "Vineland"]

MESES_ES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO",
    7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}
DIAS_ES = {0: "LUNES", 1: "MARTES", 2: "MIÉRCOLES", 3: "JUEVES", 4: "VIERNES", 5: "SÁBADO", 6: "DOMINGO"}
MES_CORTO = {1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
             7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"}


def _to_float(v):
    if v is None:
        return 0.0
    return float(str(v).replace(",", "").replace("$", "").strip() or 0)


def parse_excel(path):
    """Devuelve (fecha_date, dict_consolidado_por_sucursal) usando Net Sales (columna C)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    # Fecha del titulo: 'multi-venue - Sales Summary - 2026-06-28/2026-06-28'
    title = str(rows[0][0] or "")
    m = re.search(r"(\d{4}-\d{2}-\d{2})\s*/\s*(\d{4}-\d{2}-\d{2})", title)
    if not m:
        raise ValueError(f"No pude leer la fecha del titulo del Excel: {title!r}")
    fecha_fin = datetime.strptime(m.group(2), "%Y-%m-%d").date()

    consolidado = {b: 0.0 for b in BRANCH_ORDER}
    for r in rows[2:]:
        name = str(r[0] or "").strip()
        if not name or name.upper().startswith("REPORT"):
            continue
        if name not in VENUE_MAP:
            raise ValueError(f"Venue desconocido en {path}: {name!r}. Agregalo a VENUE_MAP.")
        consolidado[VENUE_MAP[name]] += _to_float(r[2])  # col C = Net Sales
    return fecha_fin, consolidado


def load_accumulator(path):
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(
            f"No existe {path}. Cargá el seed inicial (acumulado previo por sucursal)."
        )
    return json.loads(p.read_text(encoding="utf-8"))


def save_accumulator(path, data):
    Path(path).write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")


def money(v):
    return f"${v:,.2f}"


def build_report(ventas_path, acum25_path, acum_state_path):
    fecha, venta_dia = parse_excel(ventas_path)
    _, acum25 = parse_excel(acum25_path)

    state = load_accumulator(acum_state_path)

    # Proteccion anti doble-conteo
    if state.get("last_date") == fecha.isoformat():
        gh_out = os.environ.get("GITHUB_OUTPUT")
        if gh_out:
            with open(gh_out, "a", encoding="utf-8") as f:
                f.write("send=false\n")
        print(f"[SKIP] La fecha {fecha} ya fue procesada. No se suma de nuevo ni se envia mail.")
        sys.exit(0)

    acum_prev = state["acumulado"]  # acumulado del mes ANTES de hoy, por sucursal

    # Acum.26 = previo + venta del dia
    acum26 = {b: round(acum_prev.get(b, 0.0) + venta_dia[b], 2) for b in BRANCH_ORDER}

    rows = []
    for b in BRANCH_ORDER:
        d = venta_dia[b]
        a26 = acum26[b]
        a25 = acum25[b]
        diff = a26 - a25
        pct = (diff / a25 * 100) if a25 else 0.0
        rows.append({"branch": b, "dia": d, "a26": a26, "a25": a25, "diff": diff, "pct": pct})

    totals = {
        "dia": sum(r["dia"] for r in rows),
        "a26": sum(r["a26"] for r in rows),
        "a25": sum(r["a25"] for r in rows),
    }
    totals["diff"] = totals["a26"] - totals["a25"]
    totals["pct"] = (totals["diff"] / totals["a25"] * 100) if totals["a25"] else 0.0

    propias = {
        "dia": sum(r["dia"] for r in rows if r["branch"] in PROPIAS),
        "a26": sum(r["a26"] for r in rows if r["branch"] in PROPIAS),
        "a25": sum(r["a25"] for r in rows if r["branch"] in PROPIAS),
    }
    propias["diff"] = propias["a26"] - propias["a25"]
    propias["pct"] = (propias["diff"] / propias["a25"] * 100) if propias["a25"] else 0.0

    # Persistir el nuevo acumulado y la fecha procesada
    new_state = {"last_date": fecha.isoformat(), "acumulado": acum26}

    html = render_html(fecha, rows, totals, propias)
    return html, new_state, fecha, totals


def _pct_html(pct, diff):
    color = "#2e7d32" if pct >= 0 else "#c62828"
    sign = "+" if pct >= 0 else ""
    diff_str = f"({sign}{money(diff)})" if diff < 0 else f"(+{money(diff)})"
    # diff puede ser negativo; money ya pone el signo? No: money no pone signo de resta para negativos formateados con :,.2f -> si lo pone
    diff_str = f"({money(diff)})" if diff < 0 else f"(+{money(diff)})"
    return f'<span style="color:{color};font-weight:700;">{sign}{pct:.1f}%</span><br><span style="color:{color};font-size:12px;">{diff_str}</span>'


def render_html(fecha, rows, totals, propias):
    mes = MES_CORTO[fecha.month]
    anio_corto = str(fecha.year)[2:]
    mes_ant = MES_CORTO[fecha.month]
    anio_ant = str(fecha.year - 1)[2:]
    fecha_larga = f"{DIAS_ES[fecha.weekday()]} {fecha.day} DE {MESES_ES[fecha.month]} DE {fecha.year}"

    # Filas de la tabla detalle (variaciones en verde/rojo, resto en negro/gris)
    detalle = ""
    for r in rows:
        c = "#1a7d2e" if r["pct"] >= 0 else "#c62828"
        s = "+" if r["pct"] >= 0 else ""
        dd = f"({money(r['diff'])})" if r["diff"] < 0 else f"(+{money(r['diff'])})"
        detalle += f"""
        <tr style="border-bottom:1px solid #e6e6e6;">
          <td style="padding:14px 12px;font-weight:700;color:#000000;">{r['branch']}</td>
          <td style="padding:14px 12px;text-align:right;color:#000000;font-weight:600;">{money(r['dia'])}</td>
          <td style="padding:14px 12px;text-align:right;color:#000000;font-weight:600;">{money(r['a26'])}</td>
          <td style="padding:14px 12px;text-align:right;color:#8a8a8a;">{money(r['a25'])}</td>
          <td style="padding:14px 12px;text-align:right;"><span style="color:{c};font-weight:700;">{s}{r['pct']:.1f}%</span><br><span style="color:{c};font-size:11px;">{dd}</span></td>
        </tr>"""

    tot_color = "#1a7d2e" if totals["pct"] >= 0 else "#c62828"
    tot_sign = "+" if totals["pct"] >= 0 else ""
    tot_diff = f"({money(totals['diff'])})" if totals["diff"] < 0 else f"(+{money(totals['diff'])})"

    pr_color = "#1a7d2e" if propias["pct"] >= 0 else "#c62828"
    pr_sign = "+" if propias["pct"] >= 0 else ""
    pr_diff = f"({money(propias['diff'])})" if propias["diff"] < 0 else f"(+{money(propias['diff'])})"

    html = f"""<!DOCTYPE html>
<html lang="es">
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#f0f0f0;font-family:Arial,Helvetica,sans-serif;">
<div style="max-width:640px;margin:0 auto;background:#ffffff;border:1px solid #000000;">

  <!-- HEADER -->
  <div style="background:#000000;padding:32px;text-align:center;">
    <img src="cid:logo" alt="Lucciano's" width="200" style="display:block;margin:0 auto;max-width:200px;height:auto;">
    <div style="color:#bdbdbd;font-size:13px;letter-spacing:3px;margin-top:16px;">REPORTE DE VENTAS DIARIO</div>
    <div style="display:inline-block;background:#ffffff;color:#000000;font-weight:800;font-size:13px;letter-spacing:1px;padding:8px 16px;border-radius:4px;margin-top:18px;">
      {fecha_larga}
    </div>
  </div>

  <!-- CONSOLIDADO -->
  <div style="padding:28px 32px 8px 32px;">
    <div style="color:#8a8a8a;font-size:12px;letter-spacing:3px;border-bottom:1px solid #000000;padding-bottom:14px;">CONSOLIDADO · 6 SUCURSALES</div>

    <table width="100%" cellpadding="0" cellspacing="0" style="margin-top:20px;">
      <tr>
        <td width="33%" style="padding-right:8px;vertical-align:top;">
          <div style="background:#000000;border-radius:8px;padding:18px;">
            <div style="color:#bdbdbd;font-size:11px;letter-spacing:1px;">VENTA DEL DÍA</div>
            <div style="color:#ffffff;font-size:24px;font-weight:800;margin-top:6px;">{money(totals['dia'])}</div>
            <div style="color:#8a8a8a;font-size:11px;margin-top:4px;">6 sucursales</div>
          </div>
        </td>
        <td width="33%" style="padding:0 8px;vertical-align:top;">
          <div style="background:#000000;border-radius:8px;padding:18px;">
            <div style="color:#bdbdbd;font-size:11px;letter-spacing:1px;">ACUM. {mes.upper()}/{anio_corto}</div>
            <div style="color:#ffffff;font-size:24px;font-weight:800;margin-top:6px;">{money(totals['a26'])}</div>
            <div style="color:#8a8a8a;font-size:11px;margin-top:4px;">mes en curso</div>
          </div>
        </td>
        <td width="33%" style="padding-left:8px;vertical-align:top;">
          <div style="background:#ffffff;border:1px solid #000000;border-radius:8px;padding:18px;">
            <div style="color:#8a8a8a;font-size:11px;letter-spacing:1px;">ACUM. {mes_ant.upper()}/{anio_ant}</div>
            <div style="color:#000000;font-size:24px;font-weight:800;margin-top:6px;">{money(totals['a25'])}</div>
            <div style="color:#8a8a8a;font-size:11px;margin-top:4px;">año anterior</div>
            <div style="color:{tot_color};font-size:12px;font-weight:700;margin-top:4px;">{tot_sign}{totals['pct']:.1f}% {tot_diff}</div>
          </div>
        </td>
      </tr>
    </table>

    <!-- PROPIAS -->
    <div style="background:#ffffff;border:1px solid #000000;border-radius:8px;padding:20px;margin-top:18px;">
      <div style="color:#000000;font-size:12px;font-weight:700;letter-spacing:1px;">• PROPIAS · FLORIDA MALL · WESTON · VINELAND</div>
      <table width="100%" cellpadding="0" cellspacing="0" style="margin-top:14px;">
        <tr>
          <td style="vertical-align:top;">
            <div style="color:#8a8a8a;font-size:11px;letter-spacing:1px;">VENTA DEL DÍA</div>
            <div style="color:#000000;font-size:18px;font-weight:800;margin-top:4px;">{money(propias['dia'])}</div>
          </td>
          <td style="vertical-align:top;">
            <div style="color:#8a8a8a;font-size:11px;letter-spacing:1px;">ACUM. {mes.upper()}/{anio_corto}</div>
            <div style="color:#000000;font-size:18px;font-weight:800;margin-top:4px;">{money(propias['a26'])}</div>
          </td>
          <td style="vertical-align:top;text-align:right;">
            <div style="color:#8a8a8a;font-size:11px;letter-spacing:1px;">ACUM. {mes_ant.upper()}/{anio_ant}</div>
            <div style="color:#000000;font-size:18px;font-weight:800;margin-top:4px;">{money(propias['a25'])}</div>
            <div style="color:{pr_color};font-size:12px;font-weight:700;margin-top:2px;">{pr_sign}{propias['pct']:.1f}% {pr_diff}</div>
          </td>
        </tr>
      </table>
    </div>
  </div>

  <!-- DETALLE -->
  <div style="padding:24px 32px 36px 32px;">
    <div style="color:#8a8a8a;font-size:12px;letter-spacing:3px;">DETALLE POR SUCURSAL</div>
    <table width="100%" cellpadding="0" cellspacing="0" style="margin-top:14px;border-collapse:collapse;border:1px solid #000000;">
      <thead>
        <tr style="background:#000000;">
          <th style="padding:12px;text-align:left;color:#ffffff;font-size:11px;letter-spacing:1px;">SUCURSAL</th>
          <th style="padding:12px;text-align:right;color:#ffffff;font-size:11px;letter-spacing:1px;">DÍA</th>
          <th style="padding:12px;text-align:right;color:#ffffff;font-size:11px;letter-spacing:1px;">ACUM. {mes.upper()}/{anio_corto}</th>
          <th style="padding:12px;text-align:right;color:#ffffff;font-size:11px;letter-spacing:1px;">ACUM. {mes_ant.upper()}/{anio_ant}</th>
          <th style="padding:12px;text-align:right;color:#ffffff;font-size:11px;letter-spacing:1px;">VARIACIÓN</th>
        </tr>
      </thead>
      <tbody>{detalle}
        <tr style="border-top:2px solid #000000;background:#f7f7f7;">
          <td style="padding:16px 12px;font-weight:800;color:#000000;">TOTAL</td>
          <td style="padding:16px 12px;text-align:right;font-weight:800;color:#000000;">{money(totals['dia'])}</td>
          <td style="padding:16px 12px;text-align:right;font-weight:800;color:#000000;">{money(totals['a26'])}</td>
          <td style="padding:16px 12px;text-align:right;font-weight:800;color:#000000;">{money(totals['a25'])}</td>
          <td style="padding:16px 12px;text-align:right;"><span style="color:{tot_color};font-weight:800;">{tot_sign}{totals['pct']:.1f}%</span><br><span style="color:{tot_color};font-size:11px;">{tot_diff}</span></td>
        </tr>
      </tbody>
    </table>
  </div>

</div>
</body>
</html>"""
    return html


if __name__ == "__main__":
    base = Path(__file__).parent
    ventas = base / "Ventas_ayer.xlsx"
    acum25 = base / "Acumulado_interanual.xlsx"
    state = base / "data" / "acum_jun26.json"

    html, new_state, fecha, totals = build_report(ventas, acum25, state)

    (base / "preview.html").write_text(html, encoding="utf-8")
    save_accumulator(state, new_state)

    # Asunto del mail: "Ventas <DD/MM/YYYY> | Acum <mes>: $X (var%)"
    var_txt = f"{'+' if totals['pct'] >= 0 else ''}{totals['pct']:.1f}%"
    subject = (f"Reporte Ventas {fecha.strftime('%d/%m/%Y')} | "
               f"Día {money(totals['dia'])} · Acum {MES_CORTO[fecha.month]}/{str(fecha.year)[2:]} "
               f"{money(totals['a26'])} ({var_txt} vs {fecha.year - 1})")

    # Exponer salidas al workflow de GitHub Actions
    gh_out = os.environ.get("GITHUB_OUTPUT")
    if gh_out:
        with open(gh_out, "a", encoding="utf-8") as f:
            f.write("send=true\n")
            f.write(f"subject={subject}\n")
            f.write(f"date={fecha.isoformat()}\n")

    print(f"OK - fecha {fecha} | venta dia {money(totals['dia'])} | acum26 {money(totals['a26'])} | var {totals['pct']:.1f}%")
