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


def venta_total_dia(data_dir, dia):
    """Venta total (6 sucursales) de un dia segun el historial, o None si no esta."""
    p = Path(data_dir) / f"historico_{dia.year}.json"
    if not p.exists():
        return None
    registro = json.loads(p.read_text(encoding="utf-8")).get(dia.isoformat())
    if not registro:
        return None
    total = 0.0
    for b in BRANCH_ORDER:
        v = registro.get(b, 0)
        total += float(v.get("venta", 0)) if isinstance(v, dict) else float(v)
    return total


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
    # Mismo dia de la semana anterior (ej. viernes vs viernes): la comparacion
    # dia a dia mas justa. Si no esta en el historial, el recuadro dice "sin dato".
    sem_ant = fecha - timedelta(days=7)
    venta_sem_ant = venta_total_dia(Path(acum_state_path).parent, sem_ant)
    html = render_html(fecha, rows, totals, propias, franquicias, con_tks, sem_ant, venta_sem_ant)
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
COL_LINEA = "#d5dce8"       # separador vertical de columnas en la tabla
COL_LINEA_AZUL = "#3a5487"  # idem sobre el encabezado y la fila TOTAL


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


def _chip(p, size=12, pad="3px 8px"):
    bg = "#eaf5ec" if p >= 0 else "#fbecec"
    flecha = "▲" if p >= 0 else "▼"
    return (f'<span style="display:inline-block;background:{bg};color:{_col(p)};font-weight:700;'
            f'font-size:{size}px;padding:{pad};border-radius:20px;white-space:nowrap;">{flecha} {pct_s(p)}</span>')


def _barra(ancho_pct, color, alto=12, radio="6px", desde_derecha=False):
    """Barra como <div> con ancho en %. La app de Gmail ignora el ancho en % de
    las celdas vacias (las barras quedaban de 1px); el de un div lo respeta.
    desde_derecha=True la pega al borde derecho (barras negativas)."""
    ancho_pct = max(0.0, min(100.0, ancho_pct))
    margen = f"margin-left:{100 - ancho_pct:.1f}%;" if desde_derecha else ""
    return (f'<div style="{margen}width:{ancho_pct:.1f}%;height:{alto}px;background:{color};'
            f'border-radius:{radio};font-size:1px;line-height:1px;">&nbsp;</div>')


def _tp(venta, tickets):
    return venta / tickets if tickets else 0.0


def render_html(fecha, rows, totals, propias, franquicias, con_tks=True, sem_ant=None, venta_sem_ant=None):
    """con_tks=False: no se muestran tickets/ticket promedio del mes (falta historial)."""
    import calendar
    mes = MES_CORTO[fecha.month]
    yy = str(fecha.year)[2:]
    ya = str(fecha.year - 1)[2:]
    dias_mes = calendar.monthrange(fecha.year, fecha.month)[1]
    fecha_larga = f"{DIAS_ES[fecha.weekday()].capitalize()} {fecha.day} de {MESES_ES[fecha.month].lower()} de {fecha.year}"
    T = totals

    # ---- Bloque principal: VENTA DE AYER (lo primero que se lee)
    tp_dia = _tp(T["dia"], T["tkd"])
    dia_corto = f"{DIAS_ES[fecha.weekday()].capitalize()} {fecha.day}/{fecha.month}"

    def stat(label, valor, nota, ultima=False, size=20):
        sep = "" if ultima else f"border-right:1px solid #3a5487;"
        return f"""<td width="33%" style="padding:12px 4px;text-align:center;vertical-align:top;{sep}">
          <div style="font-size:10px;letter-spacing:1px;color:{AZUL_SUAVE};">{label}</div>
          <div style="font-size:{size}px;font-weight:800;color:#ffffff;margin-top:5px;line-height:1.2;">{valor}</div>
          <div style="font-size:11px;color:{AZUL_SUAVE};margin-top:3px;">{nota}</div></td>"""

    if con_tks:
        nota_tks = f"{T['tk26']:,} en el mes"
        nota_tp = f"{money2(_tp(T['a26'], T['tk26']))} en el mes"
    else:
        nota_tks = nota_tp = "&nbsp;"

    # Vs. mismo dia de la semana anterior
    DIA_CORTO = {0: "LUN", 1: "MAR", 2: "MIÉ", 3: "JUE", 4: "VIE", 5: "SÁB", 6: "DOM"}
    if sem_ant is not None and venta_sem_ant:
        v = (T["dia"] / venta_sem_ant - 1) * 100
        color_v = VERDE_CLARO if v >= 0 else "#ff9a9a"
        flecha = "▲" if v >= 0 else "▼"
        stat_sem_ant = stat(f"VS {DIA_CORTO[sem_ant.weekday()]} {sem_ant.day}/{sem_ant.month}",
                            f'<span style="color:{color_v};white-space:nowrap;"><span style="font-size:13px;">{flecha}</span> {pct_s(v)}</span>',
                            money_s(venta_sem_ant), ultima=True)
    else:
        lbl = f"VS {DIA_CORTO[sem_ant.weekday()]} {sem_ant.day}/{sem_ant.month}" if sem_ant else "SEMANA ANT."
        stat_sem_ant = stat(lbl, "&mdash;", "sin dato", ultima=True)

    # Reparto del dia entre propias y franquicias (barra partida)
    vp, vf = propias["dia"], franquicias["dia"]
    sp = vp / (vp + vf) * 100 if (vp + vf) else 50.0
    reparto = f"""
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:18px;">
      <tr>
        <td style="font-size:11px;color:#ffffff;padding-bottom:5px;"><span style="color:#8fd0ff;">&#9632;</span> Propias <b>{money_s(vp)}</b></td>
        <td style="font-size:11px;color:#ffffff;padding-bottom:5px;text-align:right;">Franquicias <b>{money_s(vf)}</b> <span style="color:#f2c46d;">&#9632;</span></td>
      </tr>
      <tr><td colspan="2">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>
          <td width="{sp:.1f}%" style="padding:0;"><div style="height:10px;background:#8fd0ff;border-radius:5px 0 0 5px;font-size:1px;line-height:1px;">&nbsp;</div></td>
          <td width="{100 - sp:.1f}%" style="padding:0;"><div style="height:10px;background:#f2c46d;border-radius:0 5px 5px 0;font-size:1px;line-height:1px;">&nbsp;</div></td>
        </tr></table>
      </td></tr>
    </table>"""

    ayer = f"""
    <div style="font-size:11px;letter-spacing:2px;color:{AZUL_SUAVE};">VENTA DE AYER · {dia_corto.upper()}</div>
    <div style="font-size:44px;font-weight:800;color:#ffffff;letter-spacing:-1px;margin-top:6px;">{money_s(T['dia'])}</div>
    <div style="font-size:13px;color:{AZUL_SUAVE};margin-top:2px;">6 sucursales</div>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:16px;background:#2a4a85;border-radius:10px;"><tr>
      {stat("TICKETS", f"{T['tkd']:,}", nota_tks)}
      {stat("TICKET PROM.", money2(tp_dia), nota_tp)}
      {stat_sem_ant}
    </tr></table>
    {reparto}"""

    # ---- Bloque secundario: acumulado del mes + avance (tarjeta clara)
    tope = max(T["a26"], T["a25"]) * 1.04 or 1
    acumulado = f"""
    <div style="font-size:11px;letter-spacing:2px;color:{GRIS};">ACUMULADO {MESES_ES[fecha.month]}</div>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:6px;"><tr>
      <td style="vertical-align:bottom;"><div style="font-size:28px;font-weight:800;color:{AZUL};letter-spacing:-0.5px;">{money_s(T['a26'])}</div></td>
      <td style="vertical-align:bottom;text-align:right;padding-bottom:5px;">{_chip(T['pct'], 13)}</td>
    </tr></table>
    <div style="font-size:12px;color:#5b6a85;margin-top:4px;">
      <span style="color:{_col(T['pct'])};font-weight:700;">{dif_s(T['diff'])}</span>
      vs. {mes}/{ya} ({money_s(T['a25'])}) al mismo día
    </div>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:14px;">
      <tr><td style="font-size:11px;color:{AZUL};font-weight:700;padding-bottom:4px;">{mes}/{yy}</td></tr>
      <tr><td>{_barra(T['a26'] / tope * 100, VERDE if T['pct'] >= 0 else ROJO, 10)}</td></tr>
      <tr><td style="font-size:11px;color:{GRIS};padding:7px 0 4px 0;">{mes}/{ya}</td></tr>
      <tr><td>{_barra(T['a25'] / tope * 100, '#c3cddd', 10)}</td></tr>
    </table>
    <div style="font-size:11px;color:{GRIS};margin-top:7px;">Día {fecha.day} de {dias_mes} del mes</div>"""

    # ---- Variacion del mes por sucursal (barras divergentes, de mejor a peor)
    orden = sorted(rows, key=lambda r: r["pct"], reverse=True)
    tope_pct = max(abs(r["pct"]) for r in rows) or 1

    def fila_var(r):
        w = abs(r["pct"]) / tope_pct * 100
        if r["pct"] < 0:
            izq = _barra(w, ROJO, 14, "4px 0 0 4px", desde_derecha=True)
            der = "&nbsp;"
        else:
            izq = "&nbsp;"
            der = _barra(w, VERDE, 14, "0 4px 4px 0")
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

    # 6 columnas: Sucursal | Ayer | Acum. 26 | Acum. 25 | Var. % | Dif. $
    # En el celular no entra: la tabla tiene un ancho minimo y va dentro de un
    # div con scroll horizontal (se desliza de izquierda a derecha).
    def celda(contenido, primera=False, ultima=False, color=TINTA, fw="400", bg=None, borde=True):
        pad = "11px 6px 11px 12px" if primera else ("11px 12px 11px 6px" if ultima else "11px 6px")
        return (f'<td style="padding:{pad};text-align:{"left" if primera else "right"};font-size:13px;'
                f'color:{color};font-weight:{fw};white-space:nowrap;'
                # Linea vertical entre columnas (no despues de la ultima). En la fila
                # TOTAL (borde=False, fondo azul) va en un azul mas claro.
                f'{"" if ultima else f"border-right:1px solid {COL_LINEA if borde else COL_LINEA_AZUL};"}'
                f'{f"border-bottom:1px solid {LINEA};" if borde else ""}{f"background:{bg};" if bg else ""}">{contenido}</td>')

    def fila(r, nombre, sub=False):
        nc = AZUL if sub else TINTA
        fw = "800" if sub else "400"
        bg = "#eef2f9" if sub else "#ffffff"
        return (f'<tr style="background:{bg};">'
                + celda(f'<div style="font-weight:{"800" if sub else "700"};color:{nc};">{nombre}</div>{linea_tp(r)}', primera=True)
                + celda(money_s(r["dia"]))
                + celda(money_s(r["a26"]), color=nc, fw="800" if sub else "700")
                + celda(money_s(r["a25"]), color="#6b778c", fw=fw)
                + celda(_chip(r["pct"], 11, "2px 6px"))
                + celda(dif_s(r["diff"]), ultima=True, color=_col(r["pct"]), fw="700")
                + "</tr>")

    def grupo(nombre):
        return (f'<tr><td colspan="6" style="padding:14px 14px 6px 14px;font-size:10px;font-weight:800;'
                f'letter-spacing:2px;color:{AZUL};">{nombre}</td></tr>')

    by_name = {r["branch"]: r for r in rows}
    cuerpo = (grupo("PROPIAS") + "".join(fila(by_name[b], b) for b in PROPIAS)
              + fila(propias, "Subtotal Propias", sub=True)
              + grupo("FRANQUICIAS") + "".join(fila(by_name[b], b) for b in FRANQUICIAS)
              + fila(franquicias, "Subtotal Franquicias", sub=True))

    th = lambda txt, al, pad, ultima=False: (f'<th style="padding:{pad};text-align:{al};color:#ffffff;font-size:10px;'
                                            f'letter-spacing:0.5px;font-weight:700;white-space:nowrap;'
                                            f'{"" if ultima else f"border-right:1px solid {COL_LINEA_AZUL};"}">{txt}</th>')
    titulo = lambda txt: f'<div style="font-size:11px;letter-spacing:2px;color:{GRIS};margin-bottom:8px;">{txt}</div>'

    return f"""<!DOCTYPE html>
<html lang="es">
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="color-scheme" content="light only"><meta name="supported-color-schemes" content="light">
<style>@media (min-width:620px){{ .pista-scroll{{display:none !important;}} }}</style></head>
<body style="margin:0;padding:0;background:#e9edf3;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#e9edf3;">
<tr><td align="center" style="padding:16px 8px;">

<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;max-width:600px;table-layout:fixed;background:#ffffff;border-radius:14px;overflow:hidden;font-family:Arial,Helvetica,sans-serif;">

  <!-- HEADER -->
  <tr><td style="background:{AZUL_OSC};padding:18px 22px;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>
      <td style="vertical-align:middle;"><img src="cid:logo" alt="Lucciano's" width="112" style="display:block;max-width:112px;height:auto;"></td>
      <td style="text-align:right;vertical-align:middle;">
        <div style="color:#8fa3c8;font-size:10px;letter-spacing:2px;">REPORTE DIARIO</div>
        <div style="color:#ffffff;font-size:14px;font-weight:700;margin-top:4px;">{fecha_larga}</div></td>
    </tr></table>
  </td></tr>

  <!-- VENTA DE AYER (principal) -->
  <tr><td style="padding:18px 16px 4px 16px;">
    <div style="background:{AZUL};border-radius:14px;padding:22px 20px;">{ayer}</div>
  </td></tr>

  <!-- ACUMULADO DEL MES (secundario) -->
  <tr><td style="padding:14px 16px 4px 16px;">
    <div style="background:#f3f6fb;border:1px solid {LINEA};border-radius:14px;padding:18px 20px;">{acumulado}</div>
  </td></tr>

  <!-- VARIACION POR SUCURSAL -->
  <tr><td style="padding:24px 24px 4px 24px;">
    {titulo(f"VARIACIÓN DEL MES · {yy} vs {ya}")}
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0">{variacion}</table>
  </td></tr>

  <!-- DETALLE -->
  <tr><td style="padding:24px 8px 24px 8px;">
    <div style="margin:0 12px;">{titulo("DETALLE POR SUCURSAL")}</div>
    <div class="pista-scroll" style="margin:0 12px 6px 12px;font-size:11px;color:{GRIS};">Deslizá la tabla para ver todas las columnas →</div>
    <div style="overflow-x:auto;-webkit-overflow-scrolling:touch;border:1px solid {LINEA};border-radius:12px;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;min-width:500px;">
      <tr style="background:{AZUL};">
        {th("SUCURSAL", "left", "10px 6px 10px 12px")}{th("AYER", "right", "10px 6px")}{th(f"ACUM. {mes.upper()}/{yy}", "right", "10px 6px")}{th(f"ACUM. {mes.upper()}/{ya}", "right", "10px 6px")}{th("VAR. %", "right", "10px 6px")}{th("DIF. $", "right", "10px 12px 10px 6px", ultima=True)}
      </tr>
      {cuerpo}
      <tr style="background:{AZUL};">
        {celda(f'<div style="font-weight:800;color:#ffffff;">TOTAL</div>{linea_tp(T, AZUL_SUAVE)}', primera=True, borde=False)}
        {celda(money_s(T['dia']), color="#ffffff", fw="800", borde=False)}
        {celda(money_s(T['a26']), color="#ffffff", fw="800", borde=False)}
        {celda(money_s(T['a25']), color=AZUL_SUAVE, fw="800", borde=False)}
        {celda(_chip(T['pct'], 11, "2px 6px"), borde=False)}
        {celda(dif_s(T['diff']), ultima=True, color=VERDE_CLARO if T['pct'] >= 0 else '#ff9a9a', fw="800", borde=False)}
      </tr>
    </table>
    </div>
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
