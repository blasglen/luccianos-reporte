# Lucciano's — Reporte de Ventas Diario (100% automático por GitHub Actions)

Todos los días, en forma automática, el sistema baja el reporte de ventas de
TouchBistro que llega por mail, arma la comparación contra el mismo período de
2025 y envía el reporte por Gmail. **Sin subir nada a mano.**

## Cómo funciona (flujo automático)

1. **TouchBistro** manda por mail el reporte diario (Sales Summary del día) a una
   casilla de Gmail. Sale ~2 horas después del cierre del día de servicio.
2. El workflow corre por **horario (cron)** y hace todo en un solo job:
   - `fetch_touchbistro.py` → lee esa casilla por **IMAP**, baja el adjunto y lo
     guarda como `Ventas_ayer.xlsx`. (Distingue el reporte diario del mensual por
     el rango de fechas del nombre del adjunto: agarra el de un solo día.)
   - `generar_acum2025.py` → lee el día de corte del reporte y arma
     `Acumulado_interanual.xlsx` sumando el master 2025 desde el 1° del mes hasta
     ese día, en el formato exacto que espera `report.py`.
   - `report.py` (sin cambios) → consolida las 7 filas en 6 sucursales (las dos
     Vineland se suman), toma el acumulado del mes de `data/acumulado.json`, le
     suma la venta del día, calcula la variación contra 2025 y genera el HTML.
   - `send_mail.py` (sin cambios) → envía el mail por Gmail SMTP.
   - Commitea el acumulado actualizado y los gráficos.

La **fecha del header** sale del título del propio Excel (`...2026-07-01/2026-07-01`),
así nunca se desfasa de los datos.

### El comparativo 2025 sale de un master fijo
Como 2025 ya pasó y no cambia, la venta diaria de 2025 está cargada una sola vez
en `data/Ventas_Master_2025.xlsx` (hoja "Por Dia y Local", columna `Sales` = Net
Sales). El sistema reconstruye solo el acumulado 2025 hasta cualquier día de corte.
Cubre **junio a diciembre 2025**, así que sirve para todo 2026 de junio en adelante.

### Protección anti doble-conteo
Si el mismo día se procesa dos veces, detecta que la fecha ya fue registrada
(`last_date` en el JSON) y **no vuelve a sumar ni a enviar**.

### Reinicio mensual automático
El 1° de cada mes, `report.py` detecta el cambio de mes y **reinicia el acumulado
a cero** solo. El primer día del mes, el "acumulado" es igual a la "venta del día".

## Secrets a cargar (Settings → Secrets and variables → Actions)

| Secret | Qué es |
|---|---|
| `IMAP_USER` | Casilla que **recibe** el mail de TouchBistro (para leerlo por IMAP). |
| `IMAP_APP_PASS` | App Password de 16 caracteres de **esa** casilla receptora. |
| `GMAIL_USER` | Casilla que **envía** el reporte armado. |
| `GMAIL_APP_PASS` | App Password de la casilla emisora. |
| `MAIL_TO` | Destinatario(s) del reporte. Varios: separados por coma. |

> Los App Password requieren verificación en 2 pasos activada en cada cuenta.
> Se generan en https://myaccount.google.com/apppasswords

## El horario (cron)

El workflow corre según el `cron` de `.github/workflows/reporte-ventas.yml`
(en **UTC**). Ajustalo para que dispare un rato **después** de que llega el mail
de TouchBistro. Ejemplo: si el mail cae ~2 AM US Eastern (EDT = UTC-4), eso es
~6 AM UTC → poné el cron ~7 AM UTC.

## Probar / correr a mano

Actions → "Reporte de Ventas Diario" → **Run workflow**. Si el mail más reciente
de TouchBistro es de una fecha ya procesada, los pasos corren en verde pero **no
se envía** (protección anti doble-conteo); el envío real ocurre con una fecha nueva.

## Archivos del proyecto

- `fetch_touchbistro.py` — baja el reporte del mail (IMAP).
- `generar_acum2025.py` — arma el comparativo 2025 desde el master.
- `report.py` / `send_mail.py` / `charts.py` — el motor del reporte (no se tocan).
- `data/Ventas_Master_2025.xlsx` — venta diaria 2025 por local (semilla fija).
- `data/acumulado.json` — estado del acumulado del mes en curso.
- `.github/workflows/reporte-ventas.yml` — el workflow (cron).

## Cargar meses de 2025 futuros

El master cubre jun–dic 2025. Para comparar meses de 2027 en adelante, agregá esos
meses a la hoja "Por Dia y Local" de `data/Ventas_Master_2025.xlsx` con el mismo
formato (una fila por día y local).
