# Lucciano's — Reporte de Ventas Diario (automático por GitHub Actions)

Cada vez que subís los dos Excel al repo, GitHub Actions genera el mail de ventas
(comparando el mes en curso 2026 contra el mismo período de 2025) y lo envía por Gmail.

## Cómo funciona

1. Subís (push a `main`) los archivos:
   - **`Ventas_ayer.xlsx`** → la venta del día (formato del sistema, 7 venues).
   - **`Acumulado_interanual.xlsx`** → el acumulado 2025 del mismo período corrido.
2. El workflow corre `report.py`:
   - Consolida las 7 filas en 6 sucursales (las dos *Vineland* se suman).
   - Toma el acumulado 2026 guardado en `data/acum_jun26.json` y le suma la venta del día.
   - Calcula la variación contra 2025.
   - Genera el HTML del mail.
3. Envía el mail al destinatario fijo (`send_mail.py`, vía Gmail SMTP).
4. Guarda (commit) el nuevo acumulado en `data/acum_jun26.json` para el día siguiente.

La **fecha del header** sale del título del propio Excel (`...2026-06-28/2026-06-28`),
que es la venta de "ayer". Así nunca se desfasa de los datos.

### Protección anti doble-conteo
Si subís el mismo día dos veces, el sistema detecta que la fecha ya fue procesada
(`last_date` en el JSON) y **no vuelve a sumar ni a enviar**.

## Secrets a cargar (Settings → Secrets and variables → Actions → New repository secret)

| Secret | Qué es |
|---|---|
| `GMAIL_USER` | La casilla emisora (ej. la cuenta Gmail / Google Workspace que manda). |
| `GMAIL_APP_PASS` | **App Password** de 16 caracteres (NO la contraseña normal). |
| `MAIL_TO` | El destinatario fijo del reporte. |

### Cómo generar el App Password de Gmail
1. La cuenta debe tener **verificación en 2 pasos** activada.
2. Andá a https://myaccount.google.com/apppasswords
3. Creá una app password (nombre libre, ej. "Reportes Luccianos").
4. Copiá los 16 caracteres y pegalos en el secret `GMAIL_APP_PASS` (sin espacios).

> Si la cuenta es de Google Workspace (`@luccianos.com.ar`), el admin debe permitir
> App Passwords. Si no, se puede usar una cuenta Gmail común como emisora.

## El seed inicial (ya cargado)

`data/acum_jun26.json` arranca con el acumulado **Jun 01–27 2026** que pasaste:

| Sucursal | Acum. Jun/26 (al 27) |
|---|---|
| Florida Mall | 66,132.36 |
| Weston | 59,033.56 |
| Vineland | 36,564.34 |
| American Dream | 62,482.23 |
| Sawgrass | 75,286.62 |
| Aventura | 33,196.18 |
| **Total** | **332,695.29** |

A partir del primer push (día 28), el sistema lo va actualizando solo.

## Probar localmente
```bash
pip install openpyxl
python report.py        # genera preview.html y actualiza el JSON
```

## Cambiar de mes
Cuando arranque un mes nuevo, reseteá `data/acum_jun26.json` con el acumulado en cero
(o el seed del nuevo mes) y `last_date` al último día del mes anterior. El nombre del
archivo es indistinto; si querés renombrarlo, ajustá la ruta en `report.py`.
