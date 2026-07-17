# Lucciano's USA — Reportes de Ventas

Tres reportes automáticos por mail sobre las 6 sucursales de Lucciano's USA, comparando siempre contra el mismo período del año anterior.

| Reporte | Cuándo | Período | Destinatario |
|---|---|---|---|
| **Diario** | Todos los días 7:00 | El día de ayer + acumulado del mes | `MAIL_TO` |
| **Semanal** | Lunes 7:30 | Lunes a domingo + acumulado del mes | `MAIL_TO_SOCIOS` |
| **Cierre** | Día 1 de cada mes, 8:00 | El mes completo | `MAIL_TO_SOCIOS` |

---

## Cómo se dispara

**Un solo disparador automático por reporte**: cron-job.org le pega por API a
`POST https://api.github.com/repos/blasglen/luccianos-reporte/dispatches` con un `event_type` distinto.

| Job en cron-job.org | Horario (AR) | `event_type` |
|---|---|---|
| Lucciano's - Reporte Diario | Todos los días 7:00 | `reporte-diario` |
| Lucciano's - Reporte Semanal | Lunes 7:30 | `reporte-semanal` |
| Lucciano's - Cierre Mensual | Día 1, 8:00 | `reporte-cierre` |

**No hay `schedule` en los workflows a propósito.** Dos disparadores = riesgo de mandar el mismo mail dos veces a los socios. Con uno solo eso es imposible. Si cron-job.org se cae, el reporte no sale y se corre a mano desde Actions (`workflow_dispatch` está habilitado en los tres).

El orden importa: el semanal y el cierre **leen lo que el diario dejó**, así que corren después. Si el diario falló, los dos cortan en rojo antes de mandar nada.

---

## Flujo de datos

```
                  Mail de TouchBistro (diario)
                            |
                 fetch_touchbistro.py  (IMAP)
                            |
                     Ventas_ayer.xlsx
                            |
            +---------------+---------------+
            |                               |
      historico.py                      report.py
            |                               |
 data/historico_2026.json          data/acumulado.json
 (venta + tickets por día)        (acumulado del mes; se
            |                      reinicia solo el día 2)
            |                               |
            +---------------+---------------+
                            |
        +-------------------+-------------------+
        |                   |                   |
  report.py          report_semanal.py    report_cierre.py
   (diario)              (lunes)             (día 1)
```

**Un solo punto de entrada de datos** (`fetch_touchbistro.py`), tres formas de leerlos. Si TouchBistro cambia el formato del Excel, se toca `fetch_touchbistro.py` y `VENUE_MAP`, y los tres reportes se acomodan solos.

El comparativo del año anterior sale siempre de `data/Ventas_Master_2025.xlsx` (junio a diciembre 2025, una fila por día y local), vía `generar_acum_ant.py`.

---

## Archivos

### Motores
| Archivo | Qué hace |
|---|---|
| `fetch_touchbistro.py` | Baja el adjunto diario del mail (IMAP). Distingue el diario del mensual: fecha inicio == fecha fin |
| `report.py` | Reporte diario. `parse_excel_full()` es **el** parser: consolida los 7 venues en 6 sucursales vía `VENUE_MAP` (las dos Vineland se suman) y saca Net Sales + Bill Count |
| `historico.py` | Registra el día en `data/historico_<año>.json`. Idempotente: la fecha es la clave |
| `generar_acum2025.py` | Comparativo del diario (1° del mes al día X) |
| `generar_acum_ant.py` | Comparativo **genérico por rango**. Lo usan el semanal y el cierre |
| `report_semanal.py` | Reporte semanal (lunes a domingo) |
| `report_cierre.py` | Cierre mensual + snapshot en `data/cierres.json` |
| `charts.py` | Los gráficos (matplotlib, PNG transparentes, embebidos por CID) |
| `send_mail.py` | Gmail SMTP. Primer destinatario en Para, el resto en CC |
| `fetch_historico.py` | **One-shot**: rescata del IMAP los días viejos que sigan en la casilla |

### Estado (lo escribe el bot, no tocar a mano)
| Archivo | Qué es |
|---|---|
| `data/acumulado.json` | Acumulado del mes por sucursal + `last_date`. Se reinicia solo al cambiar de mes |
| `data/historico_<año>.json` | Un día por clave: `{"venta": 1234.56, "tickets": 87}` por sucursal |
| `data/semanal.json` | Candado: la última semana enviada. Evita reenvíos |
| `data/cierres.json` | Snapshot de cada mes cerrado. **Es el candado y el archivo histórico a la vez** |

### Secrets
`IMAP_USER` / `IMAP_APP_PASS` (casilla que **recibe** de TouchBistro) · `GMAIL_USER` / `GMAIL_APP_PASS` (casilla que **envía**) · `MAIL_TO` (diario) · `MAIL_TO_SOCIOS` (semanal y cierre).

---

## Controles

El sistema prefiere **fallar en rojo antes que mandar un número mal**.

- **Anti doble-conteo (diario)**: si `last_date` es la fecha del Excel, no suma ni manda.
- **Candados (semanal y cierre)**: no se reenvía una semana o un mes ya enviado.
- **Validación de dependencia**: el semanal exige `last_date == domingo`; el cierre, `last_date == último día del mes`. Si el diario quedó atrasado, no se manda.
- **Días faltantes**: si al historial o al master 2025 les falta un día del rango, corta y **dice cuáles**.
- **Conciliación automática**: cuando el historial cubre el mes completo, el semanal compara `suma(historial)` contra `acumulado.json`. Son **dos caminos independientes al mismo número**; si no cierran al centavo, no se manda nada. Se prende sola, no hay nada que activar.
- **Venue desconocido**: si aparece una sucursal que no está en `VENUE_MAP`, revienta en vez de ignorarla.
- **Ventana del cierre**: el mes cerrado vive 24 horas en `acumulado.json` (el día 1, entre las 7:00 y las 7:00 del día 2, cuando el diario reinicia). Por eso el cierre corre el día 1 y guarda el snapshot.

---

## Criterios de negocio

- **Comparación interanual: mismas fechas calendario.** El espejo de 2026-07-13 es 2025-07-13, sin alinear por día de semana.
- **El total semanal es una comparación limpia** aunque las fechas no estén alineadas: cualquier ventana de 7 días tiene exactamente un lunes, un martes... y un domingo. La composición siempre coincide.
- **El gráfico día por día NO compara contra 2025.** Con espejo calendario, el "Lun 13" de 2026 caería al lado de un domingo de 2025. Muestra solo el año en curso con su promedio.
- **El mes SÍ tiene distorsión de calendario** (julio 2026 tiene 5 viernes; julio 2025 tenía 4). El cierre lo **blanquea con una nota automática** que aparece sola cuando las composiciones difieren.
- **Ticket promedio = venta total ÷ tickets totales.** Nunca el promedio de los promedios: eso le daría el mismo peso a Aventura que a Sawgrass.
- **Ventas netas** (Net Sales), sin impuestos. Las dos unidades de Vineland se informan consolidadas.

---

## Mantenimiento

**Sucursal nueva** → agregarla a `VENUE_MAP` (`report.py`), a `BRANCH_ORDER`, a `PROPIAS`/`FRANQUICIAS`, y a `CODE_TO_TBKEY` (`generar_acum_ant.py`). Hasta que se toque, los reportes revientan a propósito.

**Enero 2027** ⚠️ → el comparativo va a buscar el año anterior (2026) en `Ventas_Master_2025.xlsx` y no lo va a encontrar. Hay que armar `Ventas_Master_2026.xlsx` con el mismo export de TouchBistro. `historico_2026.json` cubre de julio en adelante y sirve para controlar ese export; enero a junio 2026 hay que sacarlo de TouchBistro.

**Deuda técnica conocida** → `generar_acum2025.py` duplica lógica con `generar_acum_ant.py`. Se puede convertir en un wrapper de 3 líneas que llame a `acumular_rango()`.

**Semanas que cruzan de mes** → el semanal muestra el acumulado del mes del domingo de cierre. Si la semana arranca el 31/07 y termina el 06/08, el acumulado es el de agosto. Es correcto ("acumulado del mes hasta ese día"), pero ese lunes el acumulado se ve chico contra la semana.
