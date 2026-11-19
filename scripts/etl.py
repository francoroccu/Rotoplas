"""ETL Rotoplas Perú: Google Sheet -> public/data/rotoplas.json (solo agregados).

Descarga el tablero «Ventas Rotoplas <> Patagonia» en formato xlsx y genera un JSON
con totales por mes, semana, día, asesor, familia, producto, provincia, origen y
medio de pago, más leads, cotizaciones y controles de calidad de datos.

No se publican nombres de clientes, teléfonos ni comentarios: el repositorio es público.

Uso:
  SHEET_ID=... python scripts/etl.py               # descarga desde Google Sheets
  python scripts/etl.py --file ruta/al/sheet.xlsx  # usa un archivo local
"""
import argparse
import datetime as dt
import io
import json
import os
import re
import sys
import unicodedata
import urllib.request
from collections import defaultdict
from zoneinfo import ZoneInfo

import openpyxl

LIMA = ZoneInfo("America/Lima")
OUT = os.path.join(os.path.dirname(__file__), "..", "public", "data", "rotoplas.json")
MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto",
         "Septiembre", "Octubre", "Noviembre", "Diciembre"]


def norm(s):
    s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().upper()


def num(v):
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace("S/", "").replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def to_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str):
        m = re.match(r"\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*$", v)
        if m:
            try:
                return dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            except ValueError:
                return None
    return None


def tel(v):
    """Últimos 9 dígitos de un celular peruano; None si falta o no es válido. Solo para cruces internos."""
    if v is None or v == "":
        return None
    d = re.sub(r"\D", "", str(int(v)) if isinstance(v, float) else str(v))
    return d[-9:] if len(d) >= 9 and d[-9] == "9" else "invalido"


def title(s):
    s = re.sub(r"\s+", " ", str(s or "")).strip().rstrip(".")
    return s[:1].upper() + s[1:].lower() if s else s


def place(s):
    s = re.sub(r"\s+", " ", str(s or "")).strip().rstrip(".")
    return " ".join(w.capitalize() for w in s.lower().split(" ")) if s else s


def load_workbook(path=None):
    if path:
        return openpyxl.load_workbook(path, data_only=True)
    sheet_id = os.environ.get("SHEET_ID")
    if not sheet_id:
        sys.exit("Falta la variable SHEET_ID (o usar --file).")
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    req = urllib.request.Request(url, headers={"User-Agent": "rotoplas-etl"})
    with urllib.request.urlopen(req, timeout=120) as r:
        data = r.read()
    if not data.startswith(b"PK"):
        sys.exit("La descarga no es un xlsx: revisar que la planilla siga compartida por link.")
    return openpyxl.load_workbook(io.BytesIO(data), data_only=True)


def header_row(ws, must, max_scan=15):
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan, values_only=True), start=1):
        cols = [norm(c) for c in row]
        if all(m in cols for m in must):
            return i, {c: j for j, c in enumerate(cols) if c}
    raise RuntimeError(f"No encontré encabezados {must} en {ws.title}")


def objetivos(wb):
    """Objetivo y valores de leads del tablero por período, desde Evolutivo Mensual."""
    ws = wb["Evolutivo Mensual"]
    periods, out = None, defaultdict(dict)
    keys = {"VENTA OBJETIVO (S/)": "objetivo", "LEADS GESTIONADOS": "leads_gestionados",
            "LEADS CONTACTADOS": "leads_contactados", "LEADS CONVERTIDOS": "leads_convertidos"}
    for row in ws.iter_rows(values_only=True):
        label = norm(row[0])
        if label == "INDICADOR":
            periods = [str(c) for c in row[1:] if c and re.match(r"\d{4}-\d{2}$", str(c))]
            continue
        if periods and label in keys:
            for p, v in zip(periods, row[1:1 + len(periods)]):
                out[p][keys[label]] = num(v)
    return out


def metas_asesor(wb):
    """Meta de venta por asesor y mes, del «BLOQUE 2 — VENTAS POR OPERADOR» de cada pestaña Detalle <mes>."""
    out = {}
    for ws in wb.worksheets:
        if not ws.title.startswith("Detalle") or ws.title == "Detalle de Ventas":
            continue
        per, en_bloque, metas = None, False, {}
        for r in ws.iter_rows(min_row=1, max_row=60, values_only=True):
            r = tuple(r) + (None,) * 3
            if norm(r[0]) == "PERIODO ACTIVO →".replace("→", "").strip() or str(r[0] or "").startswith("Período activo"):
                f = to_date(r[1])
                per = f"{f.year}-{f.month:02d}" if f else None
            if "VENTAS POR OPERADOR" in norm(r[0]):
                en_bloque = True
                continue
            if en_bloque:
                if norm(r[0]) in ("OPERADOR", ""):
                    continue
                if norm(r[0]) == "TOTAL" or "BLOQUE" in norm(r[0]):
                    break
                if isinstance(r[1], (int, float)) and r[1] > 0:
                    metas[place(r[0])] = round(float(r[1]), 2)
        if per and metas and per not in out:
            out[per] = metas
    return out


def ventas(wb):
    ws = wb["Detalle de Ventas"]
    h, c = header_row(ws, ["FECHA", "TOTAL", "ES VENTA"])
    rows = []
    for r in ws.iter_rows(min_row=h + 1, values_only=True):
        if r[c["FECHA"]] is None and r[c["TOTAL"]] is None:
            continue
        rows.append({
            "fecha": to_date(r[c["FECHA"]]),
            "caso": r[c["N CASO"]],
            "origen": norm(r[c["ORIGEN"]]),
            "hora": r[c["HORA"]],
            "asesor": place(r[c["ASESOR"]]) or "Sin asesor",
            "provincia": place(r[c["PROVINCIA"]]) or "Sin dato",
            "familia": norm(r[c["FAMILIA"]]) or "SIN CLASIFICAR",
            "producto": re.sub(r"\s+", " ", str(r[c["PRODUCTO"]] or "")).strip(),
            "cant": num(r[c["CANT."]]),
            "precio": num(r[c["PRECIO UNIT."]]),
            "total": num(r[c["TOTAL"]]),
            "estado": norm(r[c["ESTADO"]]),
            "canal": norm(r[c["CANAL"]]),
            "pago": norm(r[c["FUENTE / MEDIO PAGO"]]) or "SIN DATO",
            "periodo": str(r[c["PERIODO"]] or ""),
            "semana": r[c["SEMANA"]],
            "es_venta": num(r[c["ES VENTA"]]) == 1,
            "caso_unico": num(r[c["CASO UNICO"]]) if "CASO UNICO" in c else None,
            "tel": tel(r[c["NUMERO"]]) if "NUMERO" in c else None,
        })
    return rows


def leads(wb):
    ws = wb["Base de Leads"]
    h, c = header_row(ws, ["FECHA", "RESOLUCION", "VENDIO?"])
    no_contact = re.compile(r"NO CONTESTA|CASILLA|BUZON|APAGADO|NO EXISTE|EQUIVOCADO|DUPLICAD|SIN RESPUESTA|NO DISPONIBLE")
    out = defaultdict(lambda: {"gestionados": 0, "contactados": 0, "vendidos": 0, "sin_periodo": 0})
    for r in ws.iter_rows(min_row=h + 1, values_only=True):
        f = to_date(r[c["FECHA"]])
        if f is None and not r[c["ORIGEN"]]:
            continue
        p = f"{f.year}-{f.month:02d}" if f and 2025 <= f.year <= 2030 else None
        key = p or "sin-fecha"
        out[key]["gestionados"] += 1
        res = norm(r[c["RESOLUCION"]])
        if res and not no_contact.search(res):
            out[key]["contactados"] += 1
        if norm(r[c["VENDIO?"]]) in ("SI", "SÍ"):
            out[key]["vendidos"] += 1
        per = str(r[c["PERIODO"]] or "")
        if not re.match(r"\d{4}-\d{2}$", per):
            out[key]["sin_periodo"] += 1
    return out


# Tipo de contacto según la RESOLUCIÓN cargada por el asesor
TIPOS = [
    ("Habló por llamada", ("CONTESTA", "AGENDADO", "VENTA CONCRETADA")),
    ("Atendió y no le interesa", ("NO DESEA", "CORTA LA LLAMADA")),
    ("Ya compró", ("YA COMPRO",)),
    ("WhatsApp 1 a 1", ("WHATSAPP",)),
    ("Campaña WhatsApp masiva", ("CAMP. WHATSAPP",)),
    ("No contesta", ("NO CONTESTA LLAMADA", "NO CONTESTA")),
    ("Otros", ("OTROS",)),
    ("Duplicado", ("DUPLICADO",)),
]
EFECTIVOS = {"Habló por llamada", "Atendió y no le interesa", "Ya compró", "WhatsApp 1 a 1"}
ORIGENES = {"LEADS MEDIALABS-TERMAS": "Medialabs · Termas", "LEADS-TERMAS&PURIF": "Termas y purificadores",
            "GOOGLE-MIX": "Google · Mix", "GOOGLE-AGROINDUSTRIA": "Google · Agroindustria",
            "RECOMPRA": "Recompra", "LEADS FERRETERIAS": "Ferreterías", "LEADS BASE": "Base"}


def tipo_contacto(res):
    r = norm(res)
    if not r:
        return "Sin gestionar"
    for nombre, vals in TIPOS:
        if r in vals:
            return nombre
    return "Otros"


def leads_detalle(wb, ventas_rows):
    """Leads por mes: tipo de contacto, origen, asesor, venta (marcada y cruzada con ventas) y calidad."""
    ws = wb["Base de Leads"]
    h, c = header_row(ws, ["FECHA", "RESOLUCION", "VENDIO?"])
    filas = []
    for r in ws.iter_rows(min_row=h + 1, values_only=True):
        raw_f, origen, numero = r[c["FECHA"]], r[c["ORIGEN"]], r[c["NUMERO"]]
        if raw_f in (None, "") and not origen and numero in (None, "") and not r[c["RESOLUCION"]]:
            continue  # fila vacía de la plantilla
        f = to_date(raw_f)
        filas.append({
            "fecha": f, "fecha_texto": isinstance(raw_f, str),
            "periodo": f"{f.year}-{f.month:02d}" if f and 2025 <= f.year <= 2030 else "sin-fecha",
            "origen": ORIGENES.get(norm(origen), place(origen)) if origen else "Sin origen",
            "asesor": place(r[c["ASESOR"]]) or "Sin asesor",
            "tipo": tipo_contacto(r[c["RESOLUCION"]]),
            "vendio": norm(r[c["VENDIO?"]]),
            "caso_venta": r[c["N CASO VENTA"]] if "N CASO VENTA" in c else None,
            "interes": bool(str(r[c["INTERES"]] or "").strip()) if "INTERES" in c else False,
            "llamado": r[c["LLAMADO"]] if "LLAMADO" in c else None,
            "tel": tel(numero),
            "periodo_planilla": str(r[c["PERIODO"]] or ""),
        })

    # Cruce con ventas por número (interno): un lead "compró" si ese número tiene una venta desde la fecha del lead
    ventas_tel = defaultdict(list)
    for v in ventas_rows:
        if v["es_venta"] and v["tel"] and v["tel"] != "invalido":
            ventas_tel[v["tel"]].append(v)
    primer_lead = {}
    for l in filas:
        if l["tel"] and l["tel"] != "invalido":
            fl = l["fecha"] or dt.date(2000, 1, 1)
            if l["tel"] not in primer_lead or fl < primer_lead[l["tel"]][0]:
                primer_lead[l["tel"]] = (fl, l["periodo"])
    venta_cruce = defaultdict(float)
    for t, (fl, per) in primer_lead.items():
        venta_cruce[per] += sum(v["total"] for v in ventas_tel.get(t, []) if not v["fecha"] or v["fecha"] >= fl)
    for l in filas:
        fl = l["fecha"] or dt.date(2000, 1, 1)
        l["compro"] = bool(l["tel"] and any(not v["fecha"] or v["fecha"] >= fl for v in ventas_tel.get(l["tel"], [])))

    rep = defaultdict(int)
    for l in filas:
        if l["tel"] and l["tel"] != "invalido":
            rep[l["tel"]] += 1

    out = []
    for p in sorted({l["periodo"] for l in filas}):
        ls = [l for l in filas if l["periodo"] == p]

        def grupo(key, orden=None):
            g = defaultdict(lambda: {"leads": 0, "efectivos": 0, "vendidos": 0, "compraron": 0})
            for l in ls:
                k = g[key(l)]
                k["leads"] += 1; k["efectivos"] += l["tipo"] in EFECTIVOS
                k["vendidos"] += l["vendio"] == "SI"; k["compraron"] += l["compro"]
            items = [{"nombre": n, **v} for n, v in g.items()]
            if orden:
                items.sort(key=lambda x: orden.index(x["nombre"]) if x["nombre"] in orden else 99)
            else:
                items.sort(key=lambda x: -x["leads"])
            return items

        orden_tipos = [t for t, _ in TIPOS[:5]] + ["No contesta", "Sin gestionar", "Otros", "Duplicado"]
        efect = sum(1 for l in ls if l["tipo"] in EFECTIVOS)
        out.append({
            "periodo": p, "leads": len(ls),
            "gestionados": sum(1 for l in ls if l["tipo"] != "Sin gestionar"),
            "efectivos": efect,
            "vendidos": sum(1 for l in ls if l["vendio"] == "SI"),
            "compraron": sum(1 for l in ls if l["compro"]),
            "venta_cruce": round(venta_cruce.get(p, 0), 2),
            "por_tipo": grupo(lambda l: l["tipo"], orden_tipos),
            "por_origen": grupo(lambda l: l["origen"]),
            "por_asesor": grupo(lambda l: l["asesor"]),
            "calidad": {
                "sin_numero": sum(1 for l in ls if not l["tel"]),
                "numero_invalido": sum(1 for l in ls if l["tel"] == "invalido"),
                "numero_repetido": sum(1 for l in ls if l["tel"] and l["tel"] != "invalido" and rep[l["tel"]] > 1),
                "sin_origen": sum(1 for l in ls if l["origen"] == "Sin origen"),
                "sin_asesor": sum(1 for l in ls if l["asesor"] == "Sin asesor"),
                "sin_resolucion": sum(1 for l in ls if l["tipo"] == "Sin gestionar"),
                "vendio_vacio": sum(1 for l in ls if not l["vendio"]),
                "sin_interes": sum(1 for l in ls if not l["interes"]),
                "fecha_texto": sum(1 for l in ls if l["fecha_texto"]),
                "periodo_con_error": sum(1 for l in ls if not re.match(r"\d{4}-\d{2}$", l["periodo_planilla"])),
                "vendido_sin_caso": sum(1 for l in ls if l["vendio"] == "SI" and not l["caso_venta"]),
                "compro_sin_marcar": sum(1 for l in ls if l["compro"] and l["vendio"] != "SI"),
                "marcado_sin_compra": sum(1 for l in ls if l["vendio"] == "SI" and not l["compro"]),
            },
        })
    return out


def cotizaciones(wb, hoy):
    ws = wb["Seguimiento Cotizaciones"]
    h, c = header_row(ws, ["FECHA", "IMPORTE", "STATUS"])
    por_estado, por_asesor, abiertas = defaultdict(float), defaultdict(float), []
    for r in ws.iter_rows(min_row=h + 1, values_only=True):
        imp = num(r[c["IMPORTE"]])
        est = title(r[c["STATUS"]]) or "Sin estado"
        if not imp:
            continue
        por_estado[est] += imp
        asesor = title(r[c["ASESOR"]]) or "Sin asesor"
        if norm(est) == "EN SEGUIMIENTO":
            por_asesor[asesor] += imp
            ult = to_date(r[c["ULTIMA LLAMADA"]]) or to_date(r[c["FECHA"]])
            abiertas.append({"caso": str(int(r[c["NRO. CASO WISE"]])) if isinstance(r[c["NRO. CASO WISE"]], (int, float)) else str(r[c["NRO. CASO WISE"]] or "s/n"),
                             "asesor": asesor, "importe": round(imp, 2),
                             "ultimo_contacto": ult.isoformat() if ult else None,
                             "dias_sin_contacto": (hoy - ult).days if ult else None})
    abiertas.sort(key=lambda x: -x["importe"])
    cerr, perd = por_estado.get("Cerrada", 0), por_estado.get("Perdida", 0)
    return {"por_estado": {k: round(v, 2) for k, v in por_estado.items()},
            "en_seguimiento_por_asesor": {k: round(v, 2) for k, v in sorted(por_asesor.items(), key=lambda x: -x[1])},
            "tasa_cierre_monto": round(cerr / (cerr + perd), 4) if cerr + perd else None,
            "abiertas": abiertas}


def top(d, n=None):
    items = sorted(d.items(), key=lambda x: -x[1]["venta"] if isinstance(x[1], dict) else -x[1])
    return items[:n] if n else items


def build(wb):
    now = dt.datetime.now(LIMA)
    hoy = now.date()
    obj = objetivos(wb)
    metas_ase = metas_asesor(wb)
    rows = ventas(wb)
    ventas_ok = [r for r in rows if r["es_venta"]]

    by_p = defaultdict(list)
    for r in ventas_ok:
        by_p[r["periodo"]].append(r)

    meses = []
    for p in sorted(k for k in by_p if re.match(r"\d{4}-\d{2}$", k)):
        rs = by_p[p]
        y, m = map(int, p.split("-"))
        venta = sum(r["total"] for r in rs)
        casos = {r["caso"] for r in rs if r["caso"] is not None}
        sin_caso = sum(1 for r in rs if r["caso"] is None)
        n_casos = len(casos) + sin_caso
        if all(r["caso_unico"] is not None for r in rs):
            n_casos = int(sum(r["caso_unico"] for r in rs))
        uni = sum(r["cant"] for r in rs)
        o = obj.get(p, {}).get("objetivo", 0)
        agg = lambda key: defaultdict(lambda: {"venta": 0.0, "lineas": 0, "unidades": 0.0, "casos": set()})

        def group(key):
            g = defaultdict(lambda: {"venta": 0.0, "lineas": 0, "unidades": 0.0, "casos": set()})
            for r in rs:
                k = key(r)
                g[k]["venta"] += r["total"]; g[k]["lineas"] += 1; g[k]["unidades"] += r["cant"]
                g[k]["casos"].add(r["caso"] if r["caso"] is not None else id(r))
            return [{"nombre": k, "venta": round(v["venta"], 2), "lineas": v["lineas"],
                     "unidades": round(v["unidades"], 2), "casos": len(v["casos"])}
                    for k, v in sorted(g.items(), key=lambda x: -x[1]["venta"])]

        caso_monto = defaultdict(float)
        caso_info, caso_fam = {}, defaultdict(lambda: defaultdict(float))
        for r in rs:
            k = r["caso"] if r["caso"] is not None else f"sn-{id(r)}"
            caso_monto[k] += r["total"]
            caso_info[k] = r
            caso_fam[k][r["familia"]] += r["total"]
        tramos = [(0, 500, "< 500"), (500, 1000, "500–1.000"), (1000, 2000, "1.000–2.000"),
                  (2000, 5000, "2.000–5.000"), (5000, 10000, "5.000–10.000"), (10000, 1e12, "> 10.000")]
        dist = [{"tramo": lab, "casos": sum(1 for v in caso_monto.values() if lo <= v < hi),
                 "venta": round(sum(v for v in caso_monto.values() if lo <= v < hi), 2)} for lo, hi, lab in tramos]
        grandes = [{"caso": str(int(k)) if isinstance(k, (int, float)) else "s/n", "venta": round(v, 2),
                    "provincia": caso_info[k]["provincia"], "asesor": caso_info[k]["asesor"],
                    "familia": title(max(caso_fam[k].items(), key=lambda x: x[1])[0])}
                   for k, v in sorted(caso_monto.items(), key=lambda x: -x[1])[:8]]

        # Casos por día y por semana con la misma regla que el total del mes:
        # si la planilla trae «caso único» en todas las líneas, se suma esa columna
        # (un caso cuenta una sola vez aunque tenga líneas en varios días).
        usa_cu = all(r["caso_unico"] is not None for r in rs)
        dias = defaultdict(lambda: [0.0, set(), 0.0])
        for r in rs:
            if r["fecha"]:
                dias[r["fecha"]][0] += r["total"]; dias[r["fecha"]][1].add(r["caso"] if r["caso"] is not None else id(r))
                dias[r["fecha"]][2] += r["caso_unico"] or 0
        sem_cu = defaultdict(float)
        for r in rs:
            sem_cu[int(r["semana"] or 0)] += r["caso_unico"] or 0

        wa = sum(r["total"] for r in rs if "WHATSAPP" in r["canal"])
        meses.append({
            "periodo": p, "nombre": f"{MESES[m - 1]} {y}",
            "venta": round(venta, 2), "objetivo": round(o, 2),
            "cumplimiento": round(venta / o, 4) if o else None,
            "casos": n_casos, "unidades": round(uni, 2),
            "ticket": round(venta / n_casos, 2) if n_casos else 0,
            "venta_whatsapp": round(wa, 2), "venta_llamada": round(venta - wa, 2),
            "semanas": [{"semana": s["nombre"], "venta": s["venta"], "casos": int(sem_cu[s["nombre"]]) if usa_cu else s["casos"]}
                        for s in sorted(group(lambda r: int(r["semana"] or 0)), key=lambda x: x["nombre"])],
            "dias": [{"fecha": d.isoformat(), "venta": round(v[0], 2), "casos": int(v[2]) if usa_cu else len(v[1])} for d, v in sorted(dias.items())],
            "asesores": [{**a, "meta": metas_ase.get(p, {}).get(a["nombre"])} for a in group(lambda r: r["asesor"])]
                        + [{"nombre": n, "venta": 0.0, "lineas": 0, "unidades": 0.0, "casos": 0, "meta": v}
                           for n, v in metas_ase.get(p, {}).items() if n not in {r["asesor"] for r in rs}],
            "familias": group(lambda r: title(r["familia"])),
            "productos": group(lambda r: r["producto"] or "Sin dato")[:15],
            "provincias": group(lambda r: r["provincia"])[:10],
            "origen": group(lambda r: title(r["origen"]) or "Sin dato"),
            "pago": group(lambda r: title(r["pago"])),
            "tramos_ticket": dist,
            "ventas_grandes": grandes,
        })

    # Leads
    lb = leads(wb)
    leads_mes = []
    for p in sorted(set(list(lb.keys()) + [m["periodo"] for m in meses])):
        b = lb.get(p, {"gestionados": 0, "contactados": 0, "vendidos": 0, "sin_periodo": 0})
        t = obj.get(p, {})
        leads_mes.append({"periodo": p, **b,
                          "conversion_real": round(b["vendidos"] / b["contactados"], 4) if b["contactados"] else None,
                          "tablero_gestionados": t.get("leads_gestionados"), "tablero_contactados": t.get("leads_contactados"),
                          "tablero_convertidos": t.get("leads_convertidos")})

    # Calidad de datos
    total_mal = [r for r in rows if r["es_venta"] and r["cant"] and r["precio"] and abs(r["total"] - r["cant"] * r["precio"]) > 1]
    fuera = [r for r in rows if r["es_venta"] and r["fecha"] and r["periodo"] and r["periodo"] != f"{r['fecha'].year}-{r['fecha'].month:02d}"]
    calidad = {
        "ventas_sin_numero_caso": sum(1 for r in ventas_ok if r["caso"] is None),
        "lineas_total_distinto_cant_x_precio": len(total_mal),
        "monto_diferencia_total": round(sum(r["cant"] * r["precio"] - r["total"] for r in total_mal), 2),
        "lineas_periodo_distinto_fecha": len(fuera),
        "lineas_familia_otros": sum(1 for r in ventas_ok if r["familia"] in ("OTROS", "SIN CLASIFICAR")),
        "leads_sin_periodo_valido": sum(b["sin_periodo"] for b in lb.values()),
    }

    fechas = [r["fecha"] for r in ventas_ok if r["fecha"] and r["fecha"] <= hoy]
    acum_v = sum(m["venta"] for m in meses)
    acum_o = sum(m["objetivo"] for m in meses)
    return {
        "actualizado": now.isoformat(timespec="minutes"),
        "datos_hasta": max(fechas).isoformat() if fechas else None,
        "moneda": "PEN",
        "acumulado": {"venta": round(acum_v, 2), "objetivo": round(acum_o, 2),
                      "cumplimiento": round(acum_v / acum_o, 4) if acum_o else None,
                      "casos": sum(m["casos"] for m in meses), "unidades": sum(m["unidades"] for m in meses)},
        "meses": meses,
        "leads": leads_mes,
        "leads_detalle": leads_detalle(wb, rows),
        "cotizaciones": cotizaciones(wb, hoy),
        "calidad": calidad,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file")
    ap.add_argument("--out", default=OUT)
    a = ap.parse_args()
    data = build(load_workbook(a.file))
    os.makedirs(os.path.dirname(a.out), exist_ok=True)
    with open(a.out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1, default=str)
    m = data["meses"][-1] if data["meses"] else {}
    print(f"OK · {len(data['meses'])} meses · último {m.get('periodo')} S/ {m.get('venta')} · datos hasta {data['datos_hasta']}")


if __name__ == "__main__":
    main()
