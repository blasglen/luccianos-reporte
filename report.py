"""
Lucciano's - Reporte de Ventas Diario
Genera el mail HTML comparando el mes en curso (2026) contra el mismo periodo del anio anterior (2025).

Flujo:
  1. Lee Ventas_ayer.xlsx       -> venta y tickets del dia (2026)
  2. Lee data/acumulado.json    -> acumulado del mes previo (persistido en el repo)
  3. Acum.26 = acum previo + venta del dia  (se vuelve a guardar)
  4. Lee Acumulado_interanual.xlsx -> venta y tickets acumulados 2025 (comparativo)
  5. Lee data/historico_<anio>.json -> tickets acumulados del mes 2026
  6. Variacion = (acum26 - acum25) / acum25
  7. Genera el HTML del mail (diseno "A4": azul marino, sin imagenes salvo el logo)
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

# Sucursales "PROPIAS" y "FRANQUICIAS"
PROPIAS = ["Florida Mall", "Weston", "Vineland"]
FRANQUICIAS = ["American Dream", "Sawgrass", "Aventura"]

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


def parse_excel_full(path):
    """Devuelve (fecha_ini, fecha_fin, ventas_por_sucursal, tickets_por_sucursal).

    Net Sales = columna C. Bill Count = columna F (los tickets del dia).
    parse_excel() quedo como wrapper de esta para no cambiarle la firma a nadie:
    el diario la sigue llamando igual y ni se entera de que ahora tambien hay
    tickets. Un solo parser, un solo VENUE_MAP, un solo lugar donde tocar.
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    title = str(rows[0][0] or "")
    m = re.search(r"(\d{4}-\d{2}-\d{2})\s*/\s*(\d{4}-\d{2}-\d{2})", title)
    if not m:
        raise ValueError(f"No pude leer la fecha del titulo del Excel: {title!r}")
    fecha_ini = datetime.strptime(m.group(1), "%Y-%m-%d").date()
    fecha_fin = datetime.strptime(m.group(2), "%Y-%m-%d").date()

    ventas = {b: 0.0 for b in BRANCH_ORDER}
    tickets = {b: 0 for b in BRANCH_ORDER}
    for r in rows[2:]:
        name = str(r[0] or "").strip()
        if not name or name.upper().startswith("REPORT"):
            continue
        if name not in VENUE_MAP:
            raise ValueError(f"Venue desconocido en {path}: {name!r}. Agregalo a VENUE_MAP.")
        ventas[VENUE_MAP[name]] += _to_float(r[2])          # col C = Net Sales
        tickets[VENUE_MAP[name]] += int(_to_float(r[5]))    # col F = Bill Count
    return fecha_ini, fecha_fin, ventas, tickets


def parse_excel(path):
    """Devuelve (fecha_inicio, fecha_fin, dict_consolidado_por_sucursal) usando Net Sales (columna C)."""
    ini, fin, ventas, _ = parse_excel_full(path)
    return ini, fin, ventas


def load_accumulator(path):
    """Estado: {month: 'YYYY-MM', last_date: 'YYYY-MM-DD', acumulado: {sucursal: monto}}.
    Si no existe, arranca vacio (el reinicio mensual lo deja en cero al primer uso)."""
    p = Path(path)
    if not p.exists():
        return {"month": None, "last_date": None, "acumulado": {}}
    return json.loads(p.read_text(encoding="utf-8"))


def save_accumulator(path, data):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    Path(path).write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")


def money(v):
    # Sin decimales (redondeado). Formato US: $28,976
    return f"${round(v):,}"


def money2(v):
    # Con 2 decimales. Solo para el TICKET PROMEDIO, donde los centavos importan.
    return f"${v:,.2f}"


def tickets_mes(data_dir, fecha):
    """Tickets del 1ro del mes hasta `fecha` por sucursal, sumados del historial.
    Devuelve None si falta algun dia o alguno esta en formato viejo (sin tickets)."""
    p = Path(data_dir) / f"historico_{fecha.year}.json"
    if not p.exists():
        return None
    hist = json.loads(p.read_text(encoding="utf-8"))
    tks = {b: 0 for b in BRANCH_ORDER}
    for d in range(1, fecha.day + 1):
        dia = hist.get(fecha.replace(day=d).isoformat())
        if dia is None:
            print(f"[AVISO] Falta el dia {fecha.replace(day=d)} en el historial: el mail sale sin tickets del mes.")
            return None
        for b in BRANCH_ORDER:
            v = dia.get(b)
            if not isinstance(v, dict) or "tickets" not in v:
                print(f"[AVISO] El dia {fecha.replace(day=d)} no tiene tickets: el mail sale sin tickets del mes.")
                return None
            tks[b] += int(v["tickets"])
    return tks


def build_report(ventas_path, acum25_path, acum_state_path):
    fecha_ini_v, fecha, venta_dia, tks_dia = parse_excel_full(ventas_path)
    a25_ini, a25_fin, acum25, tks25 = parse_excel_full(acum25_path)

    # Validacion: ambos Excel deben referirse al mismo mes y terminar el mismo dia (comparacion espejo).
    # Ventas_ayer = un dia (fecha). Acumulado_interanual = rango del anio anterior, mismo mes, hasta el mismo dia.
    errores = []
    if a25_fin.year != fecha.year - 1:
        errores.append(
            f"El Acumulado_interanual parece NO ser del año anterior: termina en {a25_fin.year}, "
            f"se esperaba {fecha.year - 1}."
        )
    if a25_fin.month != fecha.month:
        errores.append(
            f"Los meses no coinciden: Ventas es de {MES_CORTO[fecha.month]} y "
            f"Acumulado_interanual es de {MES_CORTO[a25_fin.month]}. "
            f"Subí el acumulado {fecha.year - 1} del mismo mes."
        )
    if a25_fin.day != fecha.day:
        errores.append(
            f"Los días de corte no coinciden: Ventas es del día {fecha.day} y "
            f"el Acumulado_interanual llega hasta el día {a25_fin.day}. "
            f"Deben terminar el mismo día para comparar período espejo."
        )

    if errores:
        gh_out = os.environ.get("GITHUB_OUTPUT")
        if gh_out:
            with open(gh_out, "a", encoding="utf-8") as f:
                f.write("send=false\n")
        print("[ERROR DE VALIDACION] No se envía el mail:")
        for e in errores:
            print("  - " + e)
        sys.exit(1)

    state = load_accumulator(acum_state_path)
    mes_actual = fecha.strftime("%Y-%m")  # ej. "2026-06"

    # Proteccion anti doble-conteo (misma fecha ya procesada)
    if state.get("last_date") == fecha.isoformat():
        gh_out = os.environ.get("GITHUB_OUTPUT")
        if gh_out:
            with open(gh_out, "a", encoding="utf-8") as f:
                f.write("send=false\n")
        print(f"[SKIP] La fecha {fecha} ya fue procesada. No se suma de nuevo ni se envia mail.")
        sys.exit(0)

    # Reinicio mensual: si cambio el mes (o es la primera corrida), el acumulado arranca en cero.
    if state.get("month") != mes_actual:
        print(f"[MES NUEVO] {state.get('month')} -> {mes_actual}. Acumulado reiniciado a cero.")
        acum_prev = {b: 0.0 for b in BRANCH_ORDER}
    else:
        acum_prev = state.get("acumulado", {})

    # Acum.26 = previo (del mes en curso) + venta del dia
    acum26 = {b: round(acum_prev.get(b, 0.0) + venta_dia[b], 2) for b in BRANCH_ORDER}

    # Tickets acumulados del mes 2026: salen del historial (historico.py corre
    # antes que este script, asi que el dia de hoy ya esta). Si al historial le
    # falta algun dia del mes, los tickets del mes quedarian cortos y el ticket
    # promedio mal: en ese caso no se muestran (el mail sale igual con la venta).
    tks26 = tickets_mes(Path(acum_state_path).parent, fecha)

    rows = []
    for b in BRANCH_ORDER:
        d = venta_dia[b]
        a26 = acum26[b]
        a25 = acum25[b]
        diff = a26 - a25
        pct = (diff / a25 * 100) if a25 else 0.0
        rows.append({"branch": b, "dia": d, "a26": a26, "a25": a25, "diff": diff, "pct": pct,
                     "tkd": tks_dia[b],
                     "tk26": tks26[b] if tks26 else 0,
                     "tk25": tks25[b]})

    CAMPOS = ("dia", "a26", "a25", "tkd", "tk26", "tk25")

    def subtotal(grupo):
        s = {k: sum(r[k] for r in rows if r["branch"] in grupo) for k in CAMPOS}
        s["diff"] = s["a26"] - s["a25"]
        s["pct"] = (s["diff"] / s["a25"] * 100) if s["a25"] else 0.0
        return s

    totals = subtotal(BRANCH_ORDER)

    propias = subtotal(PROPIAS)
    franquicias = subtotal(FRANQUICIAS)

    # Persistir el nuevo acumulado, el mes activo y la fecha procesada
    new_state = {"month": mes_actual, "last_date": fecha.isoformat(), "acumulado": acum26}

    # Sin graficos PNG: el diario dibuja las barras en HTML (se ven nitidas y
    # aparecen aunque el cliente de correo bloquee imagenes). Solo va el logo.
    con_tks = tks26 is not None and totals["tk25"] > 0
    html = render_html(fecha, rows, totals, propias, franquicias, con_tks)
    return html, new_state, fecha, totals, []


# --- Diseno del mail (variante "A4": azul marino corporativo) ----------------------
# Todo con tablas e estilos inline: es lo unico que Gmail/Outlook respetan.
# Ancho fluido (max 600px): en el celular ocupa la pantalla, en la compu queda centrado.
# Las barras son celdas de tabla con ancho en %, no imagenes.

AZUL = "#1f3a6e"
AZUL_OSC = "#0f2147"
AZUL_SUAVE = "#b7c4dd"
VERDE = "#1a7d2e"
ROJO = "#c62828"
VERDE_CLARO = "#9fe0ad"   # positivos sobre fondo azul
TINTA = "#111111"
GRIS = "#8a8a8a"
LINEA = "#e3e8f0"


def money_s(v):
    """Monto sin decimales con signo menos real: -1620 -> '−$1,620'."""
    return ("−" if v < 0 else "") + f"${abs(round(v)):,}"


def dif_s(v):
    """Diferencia con signo siempre: '+$932' / '−$1,620'."""
    return ("+" if v >= 0 else "−") + f"${abs(round(v)):,}"


def pct_s(p):
    return ("+" if p >= 0 else "−") + f"{abs(p):.1f}%"


def _col(p):
    return VERDE if p >= 0 else ROJO


def _chip(p, size=12):
    bg = "#eaf5ec" if p >= 0 else "#fbecec"
    flecha = "▲" if p >= 0 else "▼"
    return (f'<span style="display:inline-block;background:{bg};color:{_col(p)};font-weight:700;'
            f'font-size:{size}px;padding:3px 8px;border-radius:20px;white-space:nowrap;">{flecha} {pct_s(p)}</span>')


def _barra(ancho_pct, color, alto=12):
    ancho_pct = max(0.0, min(100.0, ancho_pct))
    return (f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>'
            f'<td width="{ancho_pct:.1f}%" style="background:{color};height:{alto}px;line-height:{alto}px;'
            f'font-size:0;border-radius:6px;">&nbsp;</td>'
            f'<td style="font-size:0;line-height:0;">&nbsp;</td></tr></table>')


def _tp(venta, tickets):
    return venta / tickets if tickets else 0.0


def render_html(fecha, rows, totals, propias, franquicias, con_tks=True):
    """con_tks=False: no se muestran tickets/ticket promedio del mes (falta historial)."""
    import calendar
    mes = MES_CORTO[fecha.month]
    yy = str(fecha.year)[2:]
    ya = str(fecha.year - 1)[2:]
    dias_mes = calendar.monthrange(fecha.year, fecha.month)[1]
    fecha_larga = f"{DIAS_ES[fecha.weekday()].capitalize()} {fecha.day} de {MESES_ES[fecha.month].lower()} de {fecha.year}"
    T = totals

    # ---- Bloque destacado: acumulado del mes + avance
    tope = max(T["a26"], T["a25"]) * 1.04 or 1
    hero = f"""
    <div style="font-size:11px;letter-spacing:2px;color:{AZUL_SUAVE};">ACUMULADO {MESES_ES[fecha.month]} · 6 SUCURSALES</div>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:8px;"><tr>
      <td style="vertical-align:bottom;"><div style="font-size:38px;font-weight:800;color:#ffffff;letter-spacing:-1px;">{money_s(T['a26'])}</div></td>
      <td style="vertical-align:bottom;text-align:right;padding-bottom:8px;">{_chip(T['pct'], 14)}</td>
    </tr></table>
    <div style="font-size:13px;color:{AZUL_SUAVE};margin-top:6px;">
      <span style="color:{VERDE_CLARO if T['pct'] >= 0 else '#ff9a9a'};font-weight:700;">{dif_s(T['diff'])}</span>
      vs. {mes}/{ya} ({money_s(T['a25'])}) al mismo día
    </div>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:18px;">
      <tr><td style="font-size:11px;color:#ffffff;font-weight:700;padding-bottom:4px;">{mes}/{yy}</td></tr>
      <tr><td>{_barra(T['a26'] / tope * 100, '#7fd18b' if T['pct'] >= 0 else '#ef7d7d')}</td></tr>
      <tr><td style="font-size:11px;color:{AZUL_SUAVE};padding:8px 0 4px 0;">{mes}/{ya}</td></tr>
      <tr><td>{_barra(T['a25'] / tope * 100, '#4b6596')}</td></tr>
    </table>
    <div style="font-size:11px;color:{AZUL_SUAVE};margin-top:8px;">Día {fecha.day} de {dias_mes} del mes</div>"""

    # ---- KPIs de ayer
    def kpi(label, valor, nota):
        return f"""<td width="33%" style="padding:14px 6px;text-align:center;vertical-align:top;">
          <div style="font-size:10px;letter-spacing:1px;color:{GRIS};">{label}</div>
          <div style="font-size:19px;font-weight:800;color:{TINTA};margin-top:6px;">{valor}</div>
          <div style="font-size:11px;color:{GRIS};margin-top:4px;">{nota}</div></td>"""

    tp_dia = _tp(T["dia"], T["tkd"])
    if con_tks:
        tp26, tp25 = _tp(T["a26"], T["tk26"]), _tp(T["a25"], T["tk25"])
        tp_var = (tp26 / tp25 - 1) * 100 if tp25 else 0.0
        nota_tks = f"{T['tk26']:,} en el mes"
        nota_tp = f"mes {money2(tp26)} <span style=\"color:{_col(tp_var)};\">{pct_s(tp_var)}</span>"
    else:
        nota_tks = nota_tp = "&nbsp;"
    kpis = (kpi("VENTA DE AYER", money_s(T["dia"]), "6 sucursales")
            + kpi("TICKETS AYER", f"{T['tkd']:,}", nota_tks)
            + kpi("TICKET PROM.", money2(tp_dia), nota_tp))

    # ---- Variacion del mes por sucursal (barras divergentes, de mejor a peor)
    orden = sorted(rows, key=lambda r: r["pct"], reverse=True)
    tope_pct = max(abs(r["pct"]) for r in rows) or 1

    def fila_var(r):
        w = abs(r["pct"]) / tope_pct * 100
        if r["pct"] < 0:
            izq = (f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>'
                   f'<td style="font-size:0;">&nbsp;</td>'
                   f'<td width="{w:.1f}%" style="background:{ROJO};height:14px;font-size:0;border-radius:4px 0 0 4px;">&nbsp;</td>'
                   f'</tr></table>')
            der = "&nbsp;"
        else:
            izq = "&nbsp;"
            der = (f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>'
                   f'<td width="{w:.1f}%" style="background:{VERDE};height:14px;font-size:0;border-radius:0 4px 4px 0;">&nbsp;</td>'
                   f'<td style="font-size:0;">&nbsp;</td></tr></table>')
        return f"""<tr>
          <td style="padding:7px 8px 7px 0;font-size:13px;color:{TINTA};white-space:nowrap;width:118px;">{r['branch']}</td>
          <td width="30%" style="padding:7px 0;border-right:1px solid #cfcfcf;">{izq}</td>
          <td width="30%" style="padding:7px 0;">{der}</td>
          <td style="padding:7px 0 7px 10px;font-size:13px;font-weight:700;color:{_col(r['pct'])};text-align:right;white-space:nowrap;width:62px;">{pct_s(r['pct'])}</td>
        </tr>"""

    variacion = "".join(fila_var(r) for r in orden)

    # ---- Tabla de detalle
    def linea_tp(r, color=GRIS):
        if not con_tks:
            return ""
        return f'<div style="font-size:11px;color:{color};margin-top:3px;">ticket prom. {money2(_tp(r["a26"], r["tk26"]))}</div>'

    def fila(r, nombre, sub=False):
        nc = AZUL if sub else TINTA
        fw = "800" if sub else "700"
        bg = "#eef2f9" if sub else "#ffffff"
        return f"""<tr style="background:{bg};">
          <td style="padding:12px 0 12px 16px;border-bottom:1px solid {LINEA};">
            <div style="font-size:14px;font-weight:{fw};color:{nc};">{nombre}</div>{linea_tp(r)}</td>
          <td style="padding:12px 8px;text-align:right;font-size:14px;color:{TINTA};border-bottom:1px solid {LINEA};white-space:nowrap;">{money_s(r['dia'])}</td>
          <td style="padding:12px 8px;text-align:right;border-bottom:1px solid {LINEA};white-space:nowrap;">
            <div style="font-size:14px;font-weight:{fw};color:{nc};">{money_s(r['a26'])}</div>
            <div style="font-size:11px;color:{GRIS};margin-top:3px;">{money_s(r['a25'])} en {ya}</div></td>
          <td style="padding:12px 16px 12px 4px;text-align:right;border-bottom:1px solid {LINEA};white-space:nowrap;">
            {_chip(r['pct'])}<div style="font-size:11px;color:{_col(r['pct'])};margin-top:4px;">{dif_s(r['diff'])}</div></td>
        </tr>"""

    def grupo(nombre):
        return (f'<tr><td colspan="4" style="padding:14px 16px 6px 16px;font-size:10px;font-weight:800;'
                f'letter-spacing:2px;color:{AZUL};">{nombre}</td></tr>')

    by_name = {r["branch"]: r for r in rows}
    cuerpo = (grupo("PROPIAS") + "".join(fila(by_name[b], b) for b in PROPIAS)
              + fila(propias, "Subtotal Propias", sub=True)
              + grupo("FRANQUICIAS") + "".join(fila(by_name[b], b) for b in FRANQUICIAS)
              + fila(franquicias, "Subtotal Franquicias", sub=True))

    th = lambda txt, al, pad: (f'<th style="padding:{pad};text-align:{al};color:#ffffff;font-size:10px;'
                              f'letter-spacing:1px;font-weight:700;">{txt}</th>')
    titulo = lambda txt: f'<div style="font-size:11px;letter-spacing:2px;color:{GRIS};margin-bottom:8px;">{txt}</div>'

    return f"""<!DOCTYPE html>
<html lang="es">
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="color-scheme" content="light only"><meta name="supported-color-schemes" content="light"></head>
<body style="margin:0;padding:0;background:#e9edf3;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#e9edf3;">
<tr><td align="center" style="padding:16px 8px;">

<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;max-width:600px;background:#ffffff;border-radius:14px;overflow:hidden;font-family:Arial,Helvetica,sans-serif;">

  <!-- HEADER -->
  <tr><td style="background:{AZUL_OSC};padding:18px 22px;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>
      <td style="vertical-align:middle;"><img src="cid:logo" alt="Lucciano's" width="112" style="display:block;max-width:112px;height:auto;"></td>
      <td style="text-align:right;vertical-align:middle;">
        <div style="color:#8fa3c8;font-size:10px;letter-spacing:2px;">REPORTE DIARIO</div>
        <div style="color:#ffffff;font-size:14px;font-weight:700;margin-top:4px;">{fecha_larga}</div></td>
    </tr></table>
  </td></tr>

  <!-- ACUMULADO DEL MES -->
  <tr><td style="padding:18px 16px 4px 16px;">
    <div style="background:{AZUL};border-radius:14px;padding:22px 20px;">{hero}</div>
  </td></tr>

  <!-- AYER -->
  <tr><td style="padding:18px 24px 4px 24px;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f3f6fb;border-radius:12px;"><tr>{kpis}</tr></table>
  </td></tr>

  <!-- VARIACION POR SUCURSAL -->
  <tr><td style="padding:24px 24px 4px 24px;">
    {titulo(f"VARIACIÓN DEL MES · {yy} vs {ya}")}
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0">{variacion}</table>
  </td></tr>

  <!-- DETALLE -->
  <tr><td style="padding:24px 12px 24px 12px;">
    <div style="margin:0 12px;">{titulo("DETALLE POR SUCURSAL")}</div>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid {LINEA};border-radius:12px;overflow:hidden;">
      <tr style="background:{AZUL};">
        {th("SUCURSAL", "left", "11px 0 11px 16px")}{th("AYER", "right", "11px 8px")}{th(f"ACUM. {mes.upper()}", "right", "11px 8px")}{th("VAR.", "right", "11px 16px 11px 4px")}
      </tr>
      {cuerpo}
      <tr style="background:{AZUL};">
        <td style="padding:14px 0 14px 16px;"><div style="font-size:14px;font-weight:800;color:#ffffff;">TOTAL</div>{linea_tp(T, AZUL_SUAVE)}</td>
        <td style="padding:14px 8px;text-align:right;font-size:14px;font-weight:800;color:#ffffff;">{money_s(T['dia'])}</td>
        <td style="padding:14px 8px;text-align:right;">
          <div style="font-size:14px;font-weight:800;color:#ffffff;">{money_s(T['a26'])}</div>
          <div style="font-size:11px;color:{AZUL_SUAVE};margin-top:3px;">{money_s(T['a25'])} en {ya}</div></td>
        <td style="padding:14px 16px 14px 4px;text-align:right;">{_chip(T['pct'])}
          <div style="font-size:11px;color:{VERDE_CLARO if T['pct'] >= 0 else '#ff9a9a'};margin-top:4px;">{dif_s(T['diff'])}</div></td>
      </tr>
    </table>
  </td></tr>

  <!-- FOOTER -->
  <tr><td style="background:{AZUL};padding:16px 24px;text-align:center;">
    <div style="color:{AZUL_SUAVE};font-size:11px;">Lucciano's USA · Ventas netas sin impuestos · Datos al {fecha.strftime('%d/%m/%Y')}</div>
  </td></tr>

</table>
</td></tr>
</table>
</body>
</html>"""


if __name__ == "__main__":
    base = Path(__file__).parent
    ventas = base / "Ventas_ayer.xlsx"
    acum25 = base / "Acumulado_interanual.xlsx"
    state = base / "data" / "acumulado.json"

    html, new_state, fecha, totals, chart_paths = build_report(ventas, acum25, state)

    (base / "preview.html").write_text(html, encoding="utf-8")
    save_accumulator(state, new_state)

    # Asunto: "Reporte Ventas DD/MM/YYYY - Lucciano's USA"
    subject = f"Reporte Ventas {fecha.strftime('%d/%m/%Y')} - Lucciano's USA"

    # Exponer salidas al workflow de GitHub Actions
    gh_out = os.environ.get("GITHUB_OUTPUT")
    if gh_out:
        with open(gh_out, "a", encoding="utf-8") as f:
            f.write("send=true\n")
            f.write(f"subject={subject}\n")
            f.write(f"date={fecha.isoformat()}\n")

    print(f"OK - fecha {fecha} | venta dia {money(totals['dia'])} | acum26 {money(totals['a26'])} | var {totals['pct']:.1f}%")
