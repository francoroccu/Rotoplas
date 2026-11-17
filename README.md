# Rotoplas Perú · Tablero comercial

Tablero del call center y canal WhatsApp de Rotoplas Perú operado por Patagonia Digital.
Se actualiza solo, todos los días, desde la planilla «Ventas Rotoplas <> Patagonia».

```
Google Sheets ──(GitHub Actions, 06:00 y 20:00 Lima)──> public/data/rotoplas.json ──(push)──> Vercel
```

## Cómo funciona

1. **`.github/workflows/sync.yml`** corre todos los días a las 06:00 y a las 20:00 de Lima (y a mano desde *Actions › Sincronizar datos › Run workflow*).
2. **`scripts/etl.py`** descarga la planilla y genera `public/data/rotoplas.json` con **solo agregados**: totales por mes, semana, día, asesor, familia, producto, provincia, origen, medio de pago y tramo de ticket; leads por mes; cotizaciones +S/ 10.000 y controles de calidad de datos.
   No se publican nombres de clientes, teléfonos ni comentarios: el repositorio es público.
3. Si los datos cambiaron, el workflow hace commit y push. Vercel despliega cada push a `main`.
4. **`public/index.html`** es el tablero (HTML estático + Chart.js) y lee el JSON.

## Configuración

| Dónde | Variable | Para qué |
|---|---|---|
| GitHub › Settings › Secrets › Actions | `SHEET_ID` | ID de la planilla de Google Sheets (no se guarda en el código) |

La descarga usa el link de exportación de Google Sheets, así que la planilla tiene que poder leerse con el link.

## Correr local

```bash
pip install -r requirements.txt
SHEET_ID=... python scripts/etl.py          # o: python scripts/etl.py --file planilla.xlsx
python -m http.server -d public 8000
```

## Cálculos

- **¿Cómo venimos?** (arriba del tablero): venta del mes en curso del día 1 al último día con datos, contra la meta del mes y contra lo que vendió cada mes anterior en los mismos días. El ritmo esperado cuenta días hábiles de lunes a sábado (no descuenta feriados). Semáforo: *En ritmo* si el % de meta alcanza al % de días hábiles transcurridos, *Algo atrasado* desde el 80% de ese ritmo, *Atrasado* por debajo.

- **Venta**: líneas de *Detalle de Ventas* con `ES VENTA = 1`, agrupadas por `PERÍODO` y `SEMANA`.
- **Objetivo**: fila «Venta Objetivo (S/)» del *Evolutivo Mensual*.
- **Leads** (pestaña *Leads*, `leads_detalle`): *Base de Leads*, mes según `FECHA` (también si está cargada como texto dd/mm/aaaa).
  - Tipo de contacto según `RESOLUCIÓN`: habló por llamada (contesta, agendado, venta concretada), atendió y no le interesa (no desea, corta), ya compró, WhatsApp 1 a 1, campaña WhatsApp masiva, no contesta, sin gestionar, otros, duplicado.
  - Contacto efectivo = las cuatro primeras. Vendieron por gestión = `¿VENDIÓ? = SÍ` sin contar «Ya compró».
  - Compraron (cruce) = el celular del lead aparece en *Detalle de Ventas* desde la fecha del lead (el teléfono se usa solo para el cruce y no se publica).
  - Calidad: sin resolución, ¿VENDIÓ? vacío, número repetido o inválido, interés vacío, PERÍODO con error, fecha en texto, ventas no marcadas y marcadas sin venta.
- **Cotizaciones**: *Seguimiento Cotizaciones*; días sin contacto desde la última llamada.
