# Rotoplas Perú · Tablero comercial

Tablero del call center y canal WhatsApp de Rotoplas Perú operado por Patagonia Digital.
Se actualiza solo, todos los días, desde la planilla «Ventas Rotoplas <> Patagonia».

```
Google Sheets ──(GitHub Actions, diario 06:00 Lima)──> public/data/rotoplas.json ──(push)──> Vercel
```

## Cómo funciona

1. **`.github/workflows/sync.yml`** corre todos los días a las 06:00 de Lima (y a mano desde *Actions › Sincronizar datos › Run workflow*).
2. **`scripts/etl.py`** descarga la planilla y genera `public/data/rotoplas.json` con **solo agregados**: totales por mes, semana, día, asesor, familia, producto, provincia, origen, medio de pago y tramo de ticket; leads por mes; cotizaciones +S/ 10.000 y controles de calidad de datos.
   No se publican nombres de clientes, teléfonos ni comentarios: el repositorio es público.
3. Si los datos cambiaron, el workflow hace commit y push. Vercel despliega cada push a `main`.
4. **`public/index.html`** es el tablero (HTML estático + Chart.js) y lee el JSON.
5. **`middleware.js`** protege el sitio con usuario y contraseña.

## Configuración

| Dónde | Variable | Para qué |
|---|---|---|
| GitHub › Settings › Secrets › Actions | `SHEET_ID` | ID de la planilla de Google Sheets (no se guarda en el código) |
| Vercel › Settings › Environment Variables | `DASHBOARD_USER`, `DASHBOARD_PASSWORD` | Acceso al tablero |

La descarga usa el link de exportación de Google Sheets, así que la planilla tiene que poder leerse con el link.

## Correr local

```bash
pip install -r requirements.txt
SHEET_ID=... python scripts/etl.py          # o: python scripts/etl.py --file planilla.xlsx
python -m http.server -d public 8000
```

## Cálculos

- **Venta**: líneas de *Detalle de Ventas* con `ES VENTA = 1`, agrupadas por `PERÍODO` y `SEMANA`.
- **Objetivo**: fila «Venta Objetivo (S/)» del *Evolutivo Mensual*.
- **Leads reales**: *Base de Leads*; contactado = resolución distinta de no contesta, casilla, apagado, duplicado o equivocado; vendido = `¿VENDIÓ? = SÍ`.
- **Cotizaciones**: *Seguimiento Cotizaciones*; días sin contacto desde la última llamada.
