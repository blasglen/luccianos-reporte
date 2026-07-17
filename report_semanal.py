"""
Lucciano's - Reporte de Ventas SEMANAL (viernes a jueves)
--------------------------------------------------------
Se manda los VIERNES y cubre del viernes anterior al jueves (ayer).

Flujo:
  1. Calcula el rango vie..jue (por defecto, la semana que termino ayer).
  2. Lee data/historico_<anio>.json -> los 7 dias del rango, por sucursal.
     Si falta alguno, CORTA EN ROJO. No se manda una semana incompleta.
  3. Arma el comparativo del anio anterior con generar_acum_ant (fechas calendario)
     y lo deja en Acumulado_semanal_ant.xlsx (queda como evidencia auditable).
  4. Calcula variaciones, subtotales Propias/Franquicias, y la serie por dia.
  5. Genera preview_semanal.html + los 3 graficos.

DIFERENCIAS DE DISENIO CONTRA EL DIARIO (a proposito):
  - No hay acumulador incremental: la semana se recalcula entera desde el
    historial cada vez. Si un dia se corrige, el semanal lo toma solo.
  - Si hay estado, pero es solo para no MANDAR dos veces la misma semana
    (data/semanal.json). Es un candado de envio, no un acumulador de plata.
"""
import argparse
import json
import os
import sys
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

from report import (
    BRANCH_ORDER, PROPIAS, FRANQUICIAS, MESES_ES, MES_CORTO, money, parse_excel,
)
from generar_acum_ant import acumular_rango, escribir_excel, espejo

BASE = Path(__file__).parent
SALIDA_ACUM_ANT = BASE / "Acumulado_semanal_ant.xlsx"
SALIDA_ACUM_MES_ANT = BASE / "Acumulado_mensual_ant.xlsx"
ESTADO_DIARIO = BASE / "data" / "acumulado.json"
ESTADO = BASE / "data" / "semanal.json"
PREVIEW = BASE / "preview_semanal.html"

DIAS_CORTOS = {0: "Lun", 1: "Mar", 2: "Mié", 3: "Jue", 4: "Vie", 5: "Sáb", 6: "Dom"}
VIERNES = 4  # date.weekday(): lunes=0 ... viernes=4


# --- 1. Rango de la semana -------------------------------------------------------
def rango_semana(hoy=None):
    """Devuelve (viernes, jueves) de la semana a reportar.

    Regla: el reporte sale el viernes y cubre hasta AYER (jueves). O sea que
    'hasta' = hoy - 1 y 'desde' = hasta - 6.

    Si por lo que sea corre otro dia (prueba manual, respaldo que se atraso),
    retrocedo al ultimo jueves cerrado en vez de reventar. Asi el reporte siempre
    cubre una semana vie..jue completa, corra el dia que corra.
    """
    hoy = hoy or date.today()
    # Cuantos dias hay que retroceder desde 'hoy' hasta el viernes de arranque.
    # Viernes -> 7 (el viernes anterior). Sabado -> 8. Jueves -> 13.
    delta = (hoy.weekday() - VIERNES) % 7 + 7
    desde = hoy - timedelta(days=delta)
    return desde, desde + timedelta(days=6)


# --- 2. Historial ---------------------------------------------------------------
def leer_semana(desde, hasta):
    """Levanta los dias del rango del historial. Falla RUIDOSA si falta alguno."""
    path = BASE / "data" / f"historico_{hasta.year}.json"
    if not path.exists():
        raise SystemExit(
            f"[ERROR] No existe {path.name}. El historial lo escribe historico.py "
            f"en el workflow diario; todavia no corrio ninguna vez."
        )
    hist = json.loads(path.read_text(encoding="utf-8"))

    dias, faltantes = [], []
    d = desde
    while d <= hasta:
        k = d.isoformat()
        if k not in hist:
            faltantes.append(d)
        else:
            dias.append((d, hist[k]))
        d += timedelta(days=1)

    if faltantes:
        print(f"[ERROR] Al historial le faltan {len(faltantes)} dia(s) de la semana "
              f"{desde} .. {hasta}:")
        for f in faltantes:
            print(f"  - {f.isoformat()} ({DIAS_CORTOS[f.weekday()]})")
        print("Probablemente fallo el reporte diario de esos dias. Revisar Actions.")
        print("NO se manda una semana incompleta: el total quedaria subestimado y la "
              "variacion contra el anio anterior seria falsa.")
        raise SystemExit(1)

    # Si el rango cruza el 1ro de enero, el historial esta partido en dos archivos.
    # Con semanas vie..jue esto pasa una vez al anio; lo aviso en vez de mentir.
    if desde.year != hasta.year:
        raise SystemExit(
            f"[ERROR] La semana {desde}..{hasta} cruza el cambio de anio y el "
            f"historial esta separado por anio. Hay que unificar a mano esta vez."
        )
    return dias


# --- 2 bis. Acumulado del mes (viene del diario) --------------------------------
def leer_acum_mes(hasta):
    """Devuelve el acumulado del mes por sucursal, al cierre del jueves.

    NO lo recalculo: lo toma tal cual de data/acumulado.json, que es el estado que
    viene manteniendo report.py dia a dia desde el 1ro del mes. Es LA fuente del
    numero que vos ya venis mirando todos los dias, asi que el semanal y el diario
    no pueden decir cosas distintas.

    La dependencia con el diario es real, asi que la hago EXPLICITA y la valido:
    si el diario no cerro el jueves, esto corta. Prefiero eso a mandarle a los
    socios un acumulado que va un dia atrasado sin que se note.
    """
    if not ESTADO_DIARIO.exists():
        raise SystemExit(f"[ERROR] No existe {ESTADO_DIARIO.name}.")
    st = json.loads(ESTADO_DIARIO.read_text(encoding="utf-8"))

    mes_esperado = hasta.strftime("%Y-%m")
    if st.get("month") != mes_esperado:
        raise SystemExit(
            f"[ERROR] El acumulado del diario es del mes {st.get('month')} y el "
            f"cierre de esta semana es {mes_esperado}. No se manda."
        )
    if st.get("last_date") != hasta.isoformat():
        raise SystemExit(
            f"[ERROR] El acumulado del diario esta cerrado al {st.get('last_date')} "
            f"y el semanal cierra el {hasta.isoformat()}.\n"
            f"        Causa tipica: el reporte diario del jueves no corrio o fallo "
            f"(mail de TouchBistro demorado). Revisar Actions y volver a correr el "
            f"semanal despues.\n"
            f"        NO se manda: el acumulado del mes quedaria incompleto."
        )

    acum = st.get("acumulado", {})
    faltan = [b for b in BRANCH_ORDER if b not in acum]
    if faltan:
        raise SystemExit(f"[ERROR] El acumulado del diario no tiene: {', '.join(faltan)}")
    return {b: round(float(acum[b]), 2) for b in BRANCH_ORDER}


def conciliar(hasta, acum_mes):
    """CONTROL: si el historial ya cubre el mes entero hasta el jueves, la suma
    del historial TIENE que dar igual al acumulado del diario. Son dos caminos
    independientes al mismo numero.

    Hoy el historial arranca a mitad de julio, asi que este control se saltea solo
    y se va a prender automaticamente en agosto, cuando haya mes completo. No hay
    nada que activar despues.

    Devuelve un texto para el log, o None si no habia con que conciliar.
    """
    path = BASE / "data" / f"historico_{hasta.year}.json"
    if not path.exists():
        return None
    hist = json.loads(path.read_text(encoding="utf-8"))

    d = date(hasta.year, hasta.month, 1)
    suma = {b: 0.0 for b in BRANCH_ORDER}
    while d <= hasta:
        k = d.isoformat()
        if k not in hist:
            return (f"[CONCILIACION] Omitida: el historial no cubre todo "
                    f"{hasta.strftime('%m/%Y')} (falta al menos {k}). "
                    f"Se prende sola cuando haya un mes completo.")
        for b in BRANCH_ORDER:
            suma[b] += hist[k].get(b, 0.0)
        d += timedelta(days=1)

    difs = {b: round(suma[b] - acum_mes[b], 2) for b in BRANCH_ORDER
            if abs(suma[b] - acum_mes[b]) > 0.01}
    if difs:
        det = " | ".join(f"{b}: {v:+,.2f}" for b, v in difs.items())
        raise SystemExit(
            f"[ERROR] CONCILIACION FALLIDA. El historial y el acumulado del diario "
            f"no cierran:\n        {det}\n"
            f"        Alguno de los dos tiene un dia de mas, de menos, o mal sumado. "
            f"NO se manda hasta entender por que."
        )
    return (f"[CONCILIACION] OK: historial y acumulado del diario coinciden "
            f"al centavo (${sum(suma.values()):,.2f}).")


# --- 3 y 4. Calculo -------------------------------------------------------------
def construir(desde, hasta):
    dias = leer_semana(desde, hasta)

    # (a) Semana actual, por sucursal
    sem_act = {b: 0.0 for b in BRANCH_ORDER}
    for _, d in dias:
        for b in BRANCH_ORDER:
            sem_act[b] += d.get(b, 0.0)

    # (b) Acumulado del mes 2026: lo trae el diario, ya validado al jueves
    mes_act = leer_acum_mes(hasta)
    msg = conciliar(hasta, mes_act)
    if msg:
        print(msg)

    # (c) Comparativo anio anterior: escribo el Excel y lo leo con el MISMO parser
    # que usa el diario. Redundante? No: garantiza que las dos comparaciones
    # apliquen identico criterio de consolidacion, y me deja el archivo como
    # respaldo por si alguien pregunta de donde salio el numero.
    def comparativo(ini, fin, salida):
        ia, fa = espejo(ini), espejo(fin)
        acum, faltan = acumular_rango(ia, fa)
        if faltan:
            print(f"[ERROR] Al master del anio anterior le faltan dias del rango "
                  f"{ia}..{fa}:")
            for f in faltan:
                print(f"  - {f.isoformat()}")
            raise SystemExit(1)
        escribir_excel(ia, fa, acum, salida)
        _, _, consolidado = parse_excel(salida)
        return consolidado

    sem_ant = comparativo(desde, hasta, SALIDA_ACUM_ANT)
    # El acumulado del mes anterior arranca SIEMPRE el 1ro del mes del jueves.
    # Si la semana cruza fin de mes (ej. vie 31/07 a jue 06/08), el acumulado es
    # el de agosto (1 al 6), igual que hace el diario al reiniciar el 1ro.
    dia1 = date(hasta.year, hasta.month, 1)
    mes_ant = comparativo(dia1, hasta, SALIDA_ACUM_MES_ANT)

    rows = []
    for b in BRANCH_ORDER:
        s26, s25 = round(sem_act[b], 2), round(sem_ant[b], 2)
        a26, a25 = mes_act[b], round(mes_ant[b], 2)
        diff = a26 - a25
        rows.append({
            "branch": b,
            "sem26": s26, "sem25": s25,
            "sem_diff": s26 - s25, "sem_pct": ((s26 - s25) / s25 * 100) if s25 else 0.0,
            "a26": a26, "a25": a25,
            "diff": diff, "pct": (diff / a25 * 100) if a25 else 0.0,
        })

    def agregar(items):
        t = {}
        for k in ("sem26", "sem25", "a26", "a25"):
            t[k] = round(sum(r[k] for r in items), 2)
        t["sem_diff"] = t["sem26"] - t["sem25"]
        t["sem_pct"] = (t["sem_diff"] / t["sem25"] * 100) if t["sem25"] else 0.0
        t["diff"] = t["a26"] - t["a25"]
        t["pct"] = (t["diff"] / t["a25"] * 100) if t["a25"] else 0.0
        return t

    totals = agregar(rows)
    propias = agregar([r for r in rows if r["branch"] in PROPIAS])
    franquicias = agregar([r for r in rows if r["branch"] in FRANQUICIAS])

    # Serie por dia: el actual sale del historial; el anio anterior, del master
    # dia por dia (por eso acumulo rangos de un solo dia).
    serie = []
    for f, d in dias:
        f_ant = espejo(f)
        ant_cod, _ = acumular_rango(f_ant, f_ant)
        serie.append({
            "fecha": f,
            "etiqueta": f"{DIAS_CORTOS[f.weekday()]} {f.day}",
            "actual": round(sum(d.values()), 2),
            "ant": round(sum(ant_cod.values()), 2),
        })

    # El mejor y el peor dia, para el texto del mail
    mejor = max(serie, key=lambda s: s["actual"])
    peor = min(serie, key=lambda s: s["actual"])

    return rows, totals, propias, franquicias, serie, mejor, peor, dia1


# --- Candado de envio -----------------------------------------------------------
def ya_enviada(desde, hasta):
    if not ESTADO.exists():
        return False
    st = json.loads(ESTADO.read_text(encoding="utf-8"))
    return st.get("last_week") == f"{desde.isoformat()}..{hasta.isoformat()}"


def marcar_enviada(desde, hasta, totals):
    ESTADO.parent.mkdir(parents=True, exist_ok=True)
    ESTADO.write_text(json.dumps({
        "last_week": f"{desde.isoformat()}..{hasta.isoformat()}",
        "enviado": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "total": round(totals["a26"], 2),
    }, indent=2, ensure_ascii=False), encoding="utf-8")


# --- 5. HTML --------------------------------------------------------------------
def chip(pct, diff, grande=False):
    up = pct >= 0
    col = "#1a7d2e" if up else "#c62828"
    bg = "#eaf5ec" if up else "#fbecec"
    s = "+" if up else ""
    dd = f"(+{money(diff)})" if diff >= 0 else f"({money(diff)})"
    fs = "15px" if grande else "12px"
    return (f'<span style="display:inline-block;background:{bg};color:{col};'
            f'font-weight:700;font-size:{fs};padding:3px 9px;border-radius:20px;'
            f'white-space:nowrap;">{s}{pct:.1f}%</span>'
            f'<div style="color:{col};font-size:11px;margin-top:3px;">{dd}</div>')


def render_html(desde, hasta, rows, totals, propias, franquicias, serie, mejor, peor, dia1):
    anio = hasta.year
    mes = MES_CORTO[hasta.month].upper()
    a26_lbl = f"ACUM. {mes}/{str(anio)[2:]}"
    a25_lbl = f"ACUM. {mes}/{str(anio - 1)[2:]}"
    if desde.month == hasta.month:
        rango_txt = f"{desde.day} AL {hasta.day} DE {MESES_ES[hasta.month]} DE {anio}"
    else:
        rango_txt = (f"{desde.day} DE {MESES_ES[desde.month]} AL "
                     f"{hasta.day} DE {MESES_ES[hasta.month]} DE {anio}")
    esp_sem = f"{espejo(desde).strftime('%d/%m/%Y')} al {espejo(hasta).strftime('%d/%m/%Y')}"
    esp_mes = f"{espejo(dia1).strftime('%d/%m')} al {espejo(hasta).strftime('%d/%m/%Y')}"

    def fila(r, zebra):
        return f"""
        <tr style="background:{zebra};">
          <td style="padding:14px 18px;font-weight:700;color:#111111;font-size:14px;">{r['branch']}</td>
          <td style="padding:14px 12px;text-align:right;color:#111111;font-size:14px;">{money(r['sem26'])}</td>
          <td style="padding:14px 12px;text-align:right;color:#111111;font-weight:700;font-size:14px;">{money(r['a26'])}</td>
          <td style="padding:14px 12px;text-align:right;color:#9a9a9a;font-size:14px;">{money(r['a25'])}</td>
          <td style="padding:14px 18px;text-align:right;">{chip(r['pct'], r['diff'])}</td>
        </tr>"""

    def encabezado_grupo(nombre):
        return f"""
        <tr style="background:#eef1f6;">
          <td colspan="5" style="padding:9px 18px;color:#1f3a6e;font-size:10px;font-weight:800;letter-spacing:2px;">{nombre}</td>
        </tr>"""

    def fila_subtotal(nombre, s):
        return f"""
        <tr style="background:#f5f5f5;border-top:1px solid #e2e2e2;">
          <td style="padding:14px 18px;font-weight:800;color:#1f3a6e;font-size:13px;">{nombre}</td>
          <td style="padding:14px 12px;text-align:right;font-weight:800;color:#1f3a6e;font-size:13px;">{money(s['sem26'])}</td>
          <td style="padding:14px 12px;text-align:right;font-weight:800;color:#1f3a6e;font-size:13px;">{money(s['a26'])}</td>
          <td style="padding:14px 12px;text-align:right;font-weight:800;color:#7a8aa5;font-size:13px;">{money(s['a25'])}</td>
          <td style="padding:14px 18px;text-align:right;">{chip(s['pct'], s['diff'])}</td>
        </tr>"""

    by_name = {r["branch"]: r for r in rows}
    cuerpo = encabezado_grupo("PROPIAS")
    for i, b in enumerate(PROPIAS):
        cuerpo += fila(by_name[b], "#ffffff" if i % 2 == 0 else "#fafafa")
    cuerpo += fila_subtotal("Subtotal Propias", propias)
    cuerpo += encabezado_grupo("FRANQUICIAS")
    for i, b in enumerate(FRANQUICIAS):
        cuerpo += fila(by_name[b], "#ffffff" if i % 2 == 0 else "#fafafa")
    cuerpo += fila_subtotal("Subtotal Franquicias", franquicias)

    prom = totals["sem26"] / len(serie)

    return f"""<!DOCTYPE html>
<html lang="es">
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#eeeeee;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#eeeeee;">
<tr><td align="center" style="padding:24px 12px;">

<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="width:600px;max-width:600px;background:#ffffff;border-radius:14px;overflow:hidden;font-family:Arial,Helvetica,sans-serif;box-shadow:0 6px 24px rgba(0,0,0,0.12);">

  <!-- HEADER -->
  <tr><td style="background:#000000;padding:38px 32px 32px 32px;text-align:center;">
    <img src="cid:logo" alt="Lucciano's" width="190" style="display:block;margin:0 auto;max-width:190px;height:auto;">
    <div style="color:#bdbdbd;font-size:12px;letter-spacing:4px;margin-top:18px;">REPORTE DE VENTAS SEMANAL</div>
    <div style="color:#ffffff;font-size:13px;font-weight:700;letter-spacing:2px;margin-top:14px;">{rango_txt}</div>
  </td></tr>

  <!-- KPIs -->
  <tr><td style="padding:30px 32px 6px 32px;">
    <div style="color:#9a9a9a;font-size:11px;letter-spacing:3px;">CONSOLIDADO · 6 SUCURSALES</div>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:16px;">
      <tr>
        <td width="33%" style="padding-right:7px;vertical-align:top;">
          <div style="background:#111111;border-radius:12px;padding:20px;height:126px;">
            <div style="color:#9a9a9a;font-size:10px;letter-spacing:1px;">VENTA DE LA SEMANA</div>
            <div style="color:#ffffff;font-size:22px;font-weight:800;margin-top:8px;letter-spacing:-0.5px;">{money(totals['sem26'])}</div>
            <div style="margin-top:8px;">{chip(totals['sem_pct'], totals['sem_diff'])}</div>
          </div>
        </td>
        <td width="34%" style="padding:0 7px;vertical-align:top;">
          <div style="background:#111111;border-radius:12px;padding:20px;height:126px;">
            <div style="color:#9a9a9a;font-size:10px;letter-spacing:1px;">{a26_lbl}</div>
            <div style="color:#ffffff;font-size:22px;font-weight:800;margin-top:8px;letter-spacing:-0.5px;">{money(totals['a26'])}</div>
            <div style="color:#777777;font-size:11px;margin-top:6px;">del 1 al {hasta.day} de {MESES_ES[hasta.month].lower()}</div>
          </div>
        </td>
        <td width="33%" style="padding-left:7px;vertical-align:top;">
          <div style="background:#f5f5f5;border-radius:12px;padding:20px;height:126px;">
            <div style="color:#9a9a9a;font-size:10px;letter-spacing:1px;">{a25_lbl}</div>
            <div style="color:#111111;font-size:22px;font-weight:800;margin-top:8px;letter-spacing:-0.5px;">{money(totals['a25'])}</div>
            <div style="margin-top:8px;">{chip(totals['pct'], totals['diff'])}</div>
          </div>
        </td>
      </tr>
    </table>
    <div style="color:#9a9a9a;font-size:11px;margin-top:10px;">
      La variación de la izquierda es semana contra semana; la de la derecha, acumulado del mes contra {anio - 1}.
    </div>
  </td></tr>

  <!-- PROGRESO -->
  <tr><td style="padding:22px 32px 6px 32px;">
    <div style="background:#fafafa;border-radius:12px;padding:20px 22px;">
      <div style="color:#9a9a9a;font-size:11px;letter-spacing:2px;margin-bottom:6px;">AVANCE DEL MES vs AÑO ANTERIOR</div>
      <img src="cid:progreso" alt="Avance del mes" width="536" style="display:block;width:100%;max-width:536px;height:auto;">
    </div>
  </td></tr>

  <!-- VENTA POR DIA -->
  <tr><td style="padding:22px 32px 6px 32px;">
    <div style="color:#9a9a9a;font-size:11px;letter-spacing:3px;margin-bottom:12px;">VENTA POR DÍA DE LA SEMANA</div>
    <img src="cid:dias" alt="Venta por día" width="536" style="display:block;width:100%;max-width:536px;height:auto;">
    <div style="color:#777777;font-size:12px;margin-top:10px;line-height:1.6;">
      Mejor día: <b style="color:#111111;">{mejor['etiqueta']}</b> con {money(mejor['actual'])} ·
      Más flojo: <b style="color:#111111;">{peor['etiqueta']}</b> con {money(peor['actual'])} ·
      Promedio: <b style="color:#111111;">{money(prom)}</b>/día
    </div>
  </td></tr>

  <!-- COMPARATIVO POR SUCURSAL -->
  <tr><td style="padding:22px 32px 6px 32px;">
    <div style="color:#9a9a9a;font-size:11px;letter-spacing:3px;margin-bottom:12px;">SEMANA POR SUCURSAL · {str(anio)[2:]} vs {str(anio - 1)[2:]}</div>
    <img src="cid:comparativo" alt="Comparativo por sucursal" width="536" style="display:block;width:100%;max-width:536px;height:auto;">
  </td></tr>

  <!-- DETALLE -->
  <tr><td style="padding:24px 32px 36px 32px;">
    <div style="color:#9a9a9a;font-size:11px;letter-spacing:3px;margin-bottom:14px;">DETALLE POR SUCURSAL</div>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;border-radius:12px;overflow:hidden;box-shadow:0 1px 6px rgba(0,0,0,0.06);">
      <thead>
        <tr style="background:#111111;">
          <th style="padding:13px 18px;text-align:left;color:#ffffff;font-size:10px;letter-spacing:1px;font-weight:700;">SUCURSAL</th>
          <th style="padding:13px 12px;text-align:right;color:#ffffff;font-size:10px;letter-spacing:1px;font-weight:700;">SEMANA</th>
          <th style="padding:13px 12px;text-align:right;color:#ffffff;font-size:10px;letter-spacing:1px;font-weight:700;">{a26_lbl}</th>
          <th style="padding:13px 12px;text-align:right;color:#ffffff;font-size:10px;letter-spacing:1px;font-weight:700;">{a25_lbl}</th>
          <th style="padding:13px 18px;text-align:right;color:#ffffff;font-size:10px;letter-spacing:1px;font-weight:700;">VARIACIÓN</th>
        </tr>
      </thead>
      <tbody>{cuerpo}
        <tr style="background:#111111;">
          <td style="padding:17px 18px;font-weight:800;color:#ffffff;font-size:14px;">TOTAL GENERAL</td>
          <td style="padding:17px 12px;text-align:right;font-weight:800;color:#ffffff;font-size:14px;">{money(totals['sem26'])}</td>
          <td style="padding:17px 12px;text-align:right;font-weight:800;color:#ffffff;font-size:14px;">{money(totals['a26'])}</td>
          <td style="padding:17px 12px;text-align:right;font-weight:800;color:#ffffff;font-size:14px;">{money(totals['a25'])}</td>
          <td style="padding:17px 18px;text-align:right;">{chip(totals['pct'], totals['diff'])}</td>
        </tr>
      </tbody>
    </table>
    <div style="color:#9a9a9a;font-size:11px;margin-top:14px;line-height:1.6;">
      Semana del {desde.strftime('%d/%m')} al {hasta.strftime('%d/%m')}, comparada contra {esp_sem}.
      Acumulado del mes comparado contra {esp_mes} (mismas fechas calendario del año anterior).
      Ventas netas (Net Sales), sin impuestos. Las dos unidades de Vineland se informan consolidadas.
    </div>
  </td></tr>

  <!-- FOOTER -->
  <tr><td style="background:#000000;padding:20px 32px;text-align:center;">
    <div style="color:#777777;font-size:11px;letter-spacing:1px;">LUCCIANO'S USA · Reporte automático generado el {date.today().strftime('%d/%m/%Y')}</div>
  </td></tr>

</table>

</td></tr>
</table>
</body>
</html>"""


# --- Main -----------------------------------------------------------------------
def _gh_out(**kv):
    p = os.environ.get("GITHUB_OUTPUT")
    if not p:
        return
    with open(p, "a", encoding="utf-8") as f:
        for k, v in kv.items():
            f.write(f"{k}={v}\n")


def main():
    ap = argparse.ArgumentParser(description="Reporte semanal viernes a jueves.")
    ap.add_argument("--hasta", help="Jueves de cierre (YYYY-MM-DD). Default: ayer.")
    ap.add_argument("--forzar", action="store_true",
                    help="Ignora el candado y regenera aunque ya se haya enviado.")
    a = ap.parse_args()

    if a.hasta:
        hasta = datetime.strptime(a.hasta, "%Y-%m-%d").date()
        desde = hasta - timedelta(days=6)
    else:
        desde, hasta = rango_semana()

    print(f"[INFO] Semana a reportar: {desde} ({DIAS_CORTOS[desde.weekday()]}) .. "
          f"{hasta} ({DIAS_CORTOS[hasta.weekday()]})")

    if ya_enviada(desde, hasta) and not a.forzar:
        print(f"[SKIP] La semana {desde}..{hasta} ya fue enviada. No se reenvia.")
        _gh_out(send="false")
        return 0

    rows, totals, propias, franquicias, serie, mejor, peor, dia1 = construir(desde, hasta)

    # El comparativo por sucursal muestra la SEMANA (26 vs 25); el de progreso, el
    # AVANCE DEL MES (que es lo que mira el diario). Por eso al comparativo le paso
    # las claves de semana renombradas a a26/a25, que es lo que espera charts.py.
    from charts import build_charts_semanal
    rows_sem = [{"branch": r["branch"], "a26": r["sem26"], "a25": r["sem25"]} for r in rows]
    build_charts_semanal(rows_sem, totals, serie,
                         f"Semana {str(hasta.year)[2:]}", f"Semana {str(hasta.year - 1)[2:]}",
                         out_dir=str(BASE / "charts"))

    PREVIEW.write_text(
        render_html(desde, hasta, rows, totals, propias, franquicias, serie, mejor, peor, dia1),
        encoding="utf-8")
    marcar_enviada(desde, hasta, totals)

    subject = (f"Reporte Semanal {desde.strftime('%d/%m')} al {hasta.strftime('%d/%m/%Y')} "
               f"- Lucciano's USA")
    _gh_out(send="true", subject=subject, week=f"{desde.isoformat()}..{hasta.isoformat()}")

    print(f"OK - semana {money(totals['sem26'])} ({totals['sem_pct']:+.1f}%) | "
          f"acum mes {money(totals['a26'])} vs {money(totals['a25'])} "
          f"({totals['pct']:+.1f}%) | mejor dia {mejor['etiqueta']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
