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

# Region y moneda: las lee mail_base.py, que es IDENTICO en los repos de USA y
# Europa. Aca: dolares con formato US ($7,259 · $14.97 · -9.9%).
MONEDA = "USD"
REGION = "USA"

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
    # Mail (diseño "combinado", igual al de Europa): lo arma mail_diario.py.
    import mail_diario
    html = mail_diario.render(fecha, rows, totals, propias, franquicias,
                              {r["branch"]: r["tkd"] for r in rows}, Path(acum_state_path).parent.parent)
    return html, new_state, fecha, totals, []


# El HTML del mail lo arma mail_diario.py (diseño "combinado", oct/2026, igual
# en USA y Europa). La version anterior ("A4" azul marino) quedo en el historial de git.


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
