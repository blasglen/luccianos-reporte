"""
historico.py
------------
Registra la venta diaria por sucursal en data/historico_<anio>.json.

Es el insumo del reporte SEMANAL: el pipeline diario ya baja el Excel de un dia,
asi que lo unico que faltaba era que alguien lo anotara en algun lado.

Que hace:
  1. Lee Ventas_ayer.xlsx (el que dejo fetch_touchbistro.py).
  2. Reusa parse_excel() de report.py -> misma consolidacion (VENUE_MAP), mismas
     6 sucursales, mismo criterio de Net Sales. Si aparece un venue desconocido,
     revienta igual que el diario. No duplico logica a proposito: si manianа se
     agrega una sucursal, se toca UN solo lugar.
  3. Escribe {"2026-07-17": {"Florida Mall": 2542.72, ...}, ...} ordenado por fecha.

IDEMPOTENTE: la fecha es la clave del diccionario. Si el mismo dia se procesa dos
veces, se pisa el valor con el mismo numero. No hay doble-conteo posible, por eso
este script NO necesita la proteccion de last_date que si necesita el acumulador.
Por la misma razon corre ANTES de report.py: aunque report.py despues decida no
enviar (fecha repetida), el historial ya quedo guardado igual.

Falla RUIDOSA: si no encuentra el Excel o el titulo no tiene fecha, corta con error.
"""
import json
import sys
from pathlib import Path

from report import parse_excel_full, BRANCH_ORDER, VENUE_MAP

BASE = Path(__file__).parent
VENTAS = BASE / "Ventas_ayer.xlsx"


def ruta_historico(anio):
    return BASE / "data" / f"historico_{anio}.json"


def cargar(path):
    if not Path(path).exists():
        return {}
    return json.loads(Path(path).read_text(encoding="utf-8"))


def guardar(path, data):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    # Ordeno por fecha para que el diff del commit sea legible y el archivo
    # se pueda leer a ojo sin herramientas.
    ordenado = {k: data[k] for k in sorted(data)}
    Path(path).write_text(
        json.dumps(ordenado, indent=2, ensure_ascii=False), encoding="utf-8"
    )


def bruta_y_descuentos(path):
    """{sucursal: (gross_sales, discounts)} del mismo Excel diario (cols B y D).
    Los usa el mail de sales tax para mostrar Gross - Discounts = base imponible."""
    import openpyxl
    from report import _to_float
    out = {b: [0.0, 0.0] for b in BRANCH_ORDER}
    filas = list(openpyxl.load_workbook(path, data_only=True).active.iter_rows(values_only=True))
    for r in filas[2:]:
        nombre = str(r[0] or "").strip()
        if not nombre or nombre.upper().startswith("REPORT") or nombre not in VENUE_MAP:
            continue
        out[VENUE_MAP[nombre]][0] += _to_float(r[1])
        out[VENUE_MAP[nombre]][1] += _to_float(r[3])
    return out


def main():
    if not VENTAS.exists():
        print(f"[ERROR] No existe {VENTAS.name}. Corrio fetch_touchbistro.py?")
        return 1

    fecha_ini, fecha_fin, ventas, tickets = parse_excel_full(VENTAS)

    # Guarda de seguridad: este script SOLO registra dias sueltos. Si por lo que
    # sea entra el reporte mensual (rango largo), no quiero que se anote como si
    # fuese un dia y me infle el historial.
    if fecha_ini != fecha_fin:
        print(f"[ERROR] {VENTAS.name} cubre {fecha_ini}..{fecha_fin}, no es un dia "
              f"suelto. No lo registro en el historial.")
        return 1

    path = ruta_historico(fecha_fin.year)
    hist = cargar(path)
    clave = fecha_fin.isoformat()
    ya_estaba = clave in hist

    # Formato: {"Sucursal": {"venta": 1234.56, "tickets": 87, "bruta": 1236.0, "descuentos": 1.44}}
    # (bruta y descuentos desde oct/2026; los dias viejos solo tienen venta y tickets)
    # Los tickets salen de Bill Count. Con venta + tickets se puede calcular el
    # ticket promedio, que es lo que contesta si vendimos mas porque vino mas
    # gente o porque cada uno gasta mas.
    bd = bruta_y_descuentos(VENTAS)
    hist[clave] = {b: {"venta": round(ventas[b], 2), "tickets": tickets[b],
                       "bruta": round(bd[b][0], 2), "descuentos": round(bd[b][1], 2)}
                   for b in BRANCH_ORDER}
    guardar(path, hist)

    total = sum(v["venta"] for v in hist[clave].values())
    tks = sum(v["tickets"] for v in hist[clave].values())
    estado = "actualizado (ya existia)" if ya_estaba else "agregado"
    print(f"[OK] Historial {estado}: {clave} = ${total:,.2f} | {tks} tickets | "
          f"{len(hist)} dias registrados en {path.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
