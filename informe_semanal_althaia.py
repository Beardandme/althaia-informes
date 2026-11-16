#!/usr/bin/env python3
"""Informe semanal del taproom Althaia (Altea): PDF + Excel/CSV importable en Odoo + texto del email.

Uso:  python3 informe_semanal.py data.json carpeta_salida

data.json (claves cortas, copiadas tal cual de las consultas a Square):
  ini, fin            fechas ISO del periodo pedido (lunes y domingo)
  dias   [{d, src, gross, ret, net, tax, tk, ref, n}]       Sales por reporting_day x order_source
  pagos  [{d, m, pay, npay, ref, nref}]                     PaymentMethods por local_date x payment_method
  horas  [{d, h, gross, ret, tk}]                           Sales por reporting_day x local_hour
  prod   [{cat, item, var, q, gross}]                       ProductMixReport por categoria x item x variacion
  sup    [{d, h, q, gross}]                                 'Suplemento concierto' por dia x hora
Salida: PDF, XLSX, CSV, email.txt y meta.json (ok=false si algo no cuadra: NO enviar sin revisar).
"""
import csv
import datetime as dt
import html
import json
import os
import shutil
import subprocess
import sys
from collections import defaultdict

DIAS = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]
MES_CORTO = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
CUENTAS = {"diario": "Operaciones varias", "efectivo": "570000", "datafono": "572000",
           "ventas": "700000", "iva": "477000"}
UMBRAL_HORARIO = 10          # un dia con >= 10 tickets lleva su grafico hora a hora
TOL = 0.10                   # tolerancia (EUR) en las comprobaciones cruzadas

CSS = """
:root { color-scheme: light; --bg:#F2F4F8; --card:#FFFFFF; --ink:#12234A; --muted:#5B6786; --line:#D9DEEA;
  --accent:#E8641A; --ok:#1E7B4F; --warn:#B7791F; --bad:#B42318;
  --display:'Anton','Impact',sans-serif; --body:'Signika','Segoe UI',system-ui,sans-serif;
  --mono:'Space Mono',ui-monospace,'Courier New',monospace; }
* { box-sizing:border-box; }
html, body { margin:0; background:#F2F4F8; }
body { background:var(--bg); color:var(--ink); font-family:var(--body); font-size:16px; line-height:1.5; }
.wrap { max-width:860px; margin:0 auto; padding:24px 16px 56px; display:flex; flex-direction:column; gap:28px; }
header { display:flex; flex-direction:column; gap:6px; border-bottom:3px solid var(--ink); padding-bottom:16px; }
.eyebrow { font-family:var(--mono); font-size:12px; letter-spacing:.08em; text-transform:uppercase; color:var(--accent); font-weight:700; }
h1 { font-family:var(--display); font-weight:400; font-size:40px; line-height:1.05; margin:0; text-transform:uppercase; text-wrap:balance; }
.sub { color:var(--muted); font-size:14px; margin:0; }
h2 { font-family:var(--display); font-weight:400; font-size:24px; text-transform:uppercase; margin:0; letter-spacing:.01em; text-wrap:balance; }
section { display:flex; flex-direction:column; gap:12px; }
.lead { margin:0; color:var(--muted); max-width:65ch; }
.kpis { display:grid; grid-template-columns:repeat(auto-fit,minmax(160px,1fr)); gap:12px; }
.kpi { background:var(--card); border:1px solid var(--line); border-radius:8px; padding:14px 14px 12px; display:flex; flex-direction:column; gap:2px; }
.kpi.main { background:var(--ink); color:#fff; border-color:var(--ink); }
.kpi .l { font-family:var(--mono); font-size:11px; text-transform:uppercase; letter-spacing:.06em; color:var(--muted); }
.kpi.main .l { color:#B9C4E3; }
.kpi .v { font-family:var(--display); font-size:32px; line-height:1.1; font-variant-numeric:tabular-nums; }
.kpi .n { font-size:13px; color:var(--muted); }
.kpi.main .n { color:#B9C4E3; }
.card { background:var(--card); border:1px solid var(--line); border-radius:8px; padding:4px 0; }
.scroll { overflow-x:auto; }
table { border-collapse:collapse; width:100%; font-size:15px; font-variant-numeric:tabular-nums; }
th, td { padding:10px 12px; text-align:right; white-space:nowrap; border-bottom:1px solid var(--line); }
th:first-child, td:first-child { text-align:left; }
td:first-child { white-space:normal; }
th { font-family:var(--mono); font-size:11px; text-transform:uppercase; letter-spacing:.05em; color:var(--muted); font-weight:400; }
tr:last-child td { border-bottom:0; }
tr.total td { font-weight:600; background:#EEF1F8; }
.dim { color:var(--muted); }
.bars { display:flex; flex-direction:column; gap:8px; padding:14px 14px; }
.row { display:grid; grid-template-columns:132px 1fr auto; gap:10px; align-items:center; font-size:14px; }
.row .name { line-height:1.2; overflow-wrap:anywhere; }
.track { background:#E7EBF3; border-radius:3px; height:14px; overflow:hidden; }
.fill { display:block; height:100%; background:var(--ink); border-radius:3px; }
.fill.o { background:var(--accent); }
.val { font-family:var(--mono); font-size:12px; text-align:right; font-variant-numeric:tabular-nums; min-width:96px; }
.pill { display:inline-block; font-family:var(--mono); font-size:11px; text-transform:uppercase; letter-spacing:.05em; padding:2px 8px; border-radius:999px; border:1px solid currentColor; vertical-align:middle; }
.pill.warn { color:var(--warn); } .pill.ok { color:var(--ok); } .pill.bad { color:var(--bad); }
.note { background:var(--card); border:1px solid var(--line); border-left:4px solid var(--accent); border-radius:6px; padding:12px 14px; font-size:15px; max-width:70ch; }
.note.bad { border-left-color:var(--bad); }
.note p { margin:0; } .note p + p { margin-top:8px; }
ul.plain { margin:0; padding-left:20px; display:flex; flex-direction:column; gap:8px; max-width:70ch; }
.small { font-size:13px; color:var(--muted); margin:0; }
footer { border-top:1px solid var(--line); padding-top:16px; }
@media print {
  * { -webkit-print-color-adjust:exact; print-color-adjust:exact; }
  section, .card, .kpi, .note, tr { break-inside:avoid; }
  h2 { break-after:avoid; }
}
@media (max-width:480px) { h1 { font-size:34px; } .row { grid-template-columns:96px 1fr auto; } .val { min-width:80px; } }
"""


# ---------------------------------------------------------------- utilidades
def fmt(x, d=2):
    """Numero en formato espanol: 1.234,56 (signo menos tipografico)."""
    s = f"{abs(x):,.{d}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("−" + s) if x < -0.5 * 10 ** (-d) else s


def pct(x, d=1):
    return fmt(x, d) + " %"


def esc(s):
    return html.escape(str(s), quote=True)


def lista_y(items):
    items = list(items)
    if len(items) <= 1:
        return "".join(items)
    return ", ".join(items[:-1]) + " y " + items[-1]


def dia_largo(d):
    x = dt.date.fromisoformat(d)
    return f"{DIAS[x.weekday()].capitalize()} {x.day}"


def dia_nombre(d):
    return DIAS[dt.date.fromisoformat(d).weekday()]


def fecha_corta(d):
    x = dt.date.fromisoformat(d)
    return x.strftime("%d/%m/%Y")


def etiqueta_periodo(days):
    ds = [dt.date.fromisoformat(d) for d in days]
    a, b = ds[0], ds[-1]

    def dn(x):
        return f"{DIAS[x.weekday()]} {x.day}"
    if a == b:
        return f"{dn(a).capitalize()} de {MESES[a.month - 1]}"
    if a.month == b.month:
        if len(ds) == 2 and (b - a).days == 1:
            return f"{dn(a).capitalize()} y {dn(b)} de {MESES[a.month - 1]}"
        return f"Del {dn(a)} al {dn(b)} de {MESES[a.month - 1]}"
    return f"Del {dn(a)} de {MESES[a.month - 1]} al {dn(b)} de {MESES[b.month - 1]}"


def nombre_archivo_periodo(days):
    a, b = dt.date.fromisoformat(days[0]), dt.date.fromisoformat(days[-1])
    if a == b:
        return f"{a.day}-{MES_CORTO[a.month - 1]}"
    if a.month == b.month:
        return f"{a.day}-{b.day}-{MES_CORTO[a.month - 1]}"
    return f"{a.day}-{MES_CORTO[a.month - 1]}_{b.day}-{MES_CORTO[b.month - 1]}"


def limpia_cerveza(n):
    return n.replace("Mediterranean ", "").strip()


# ---------------------------------------------------------------- calculo
def calcular(D):
    flags = []  # {nivel: info|error, texto}

    def flag(nivel, texto):
        flags.append({"nivel": nivel, "texto": texto})

    dias = defaultdict(lambda: defaultdict(float))
    por_src = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    for r in D.get("dias", []):
        d = r["d"][:10]
        s = r.get("src") or ""
        for k in ("gross", "ret", "net", "tax", "tk", "ref", "n"):
            dias[d][k] += float(r.get(k) or 0)
            por_src[d][s][k] += float(r.get(k) or 0)
    activos = sorted(d for d, v in dias.items() if v["tk"] > 0 or abs(v["net"]) > 0.004)
    if not activos:
        return None

    pagos = defaultdict(lambda: defaultdict(float))
    for r in D.get("pagos", []):
        d = r["d"][:10]
        pagos[d]["pay"] += float(r.get("pay") or 0)
        pagos[d]["ref"] += float(r.get("ref") or 0)
        if r.get("m") == "CASH":
            pagos[d]["cash_pay"] += float(r.get("pay") or 0)
            pagos[d]["cash_ref"] += float(r.get("ref") or 0)

    horas = defaultdict(dict)
    for r in D.get("horas", []):
        d = r["d"][:10]
        h = int(r["h"])
        x = horas[d].setdefault(h, {"gross": 0.0, "ret": 0.0, "tk": 0.0})
        x["gross"] += float(r.get("gross") or 0)
        x["ret"] += float(r.get("ret") or 0)
        x["tk"] += float(r.get("tk") or 0)

    sup = defaultdict(lambda: defaultdict(lambda: [0.0, 0.0]))
    for r in D.get("sup", []):
        if float(r.get("q") or 0) > 0:
            x = sup[r["d"][:10]][int(r["h"])]
            x[0] += float(r["q"])
            x[1] += float(r.get("gross") or 0)

    V = {}
    for d in activos:
        v = dias[d]
        base = round(v["net"], 2)
        iva = round(v["tax"], 2)
        total = round(base + iva, 2)
        efectivo = round(pagos[d]["cash_pay"] + pagos[d]["cash_ref"], 2)
        if efectivo < -0.004:
            flag("error", f"{dia_largo(d)}: el efectivo sale negativo ({fmt(efectivo)} €). Revisar devoluciones en caja.")
        datafono = round(total - efectivo, 2)
        cobrado = round(pagos[d]["pay"] + pagos[d]["ref"], 2)
        esperado = round(total + v["ref"], 2)
        if abs(cobrado - esperado) > TOL:
            flag("error", f"{dia_largo(d)}: los cobros de Square ({fmt(cobrado)} €) no cuadran con las ventas ({fmt(esperado)} €).")
        V[d] = dict(gross=round(v["gross"], 2), ret=round(v["ret"], 2), base=base, iva=iva, total=total,
                    tk=int(round(v["tk"])), ref=round(v["ref"], 2), efectivo=efectivo, datafono=datafono,
                    cobrado=cobrado)

    T = {k: round(sum(V[d][k] for d in activos), 2) for k in
         ("gross", "ret", "base", "iva", "total", "ref", "efectivo", "datafono", "cobrado")}
    T["tk"] = sum(V[d]["tk"] for d in activos)
    rate = T["iva"] / T["base"] if T["base"] > 0 else 0.10
    rate_txt = "10" if abs(rate - 0.10) < 0.005 else fmt(rate * 100, 1)
    mult = 1 + rate

    # --- producto
    def grupo(cat, item):
        if item == "Suplemento concierto":
            return "Suplemento concierto"
        if item == "Tabla de degustación":
            return "Tabla degustación"
        m = {"Cerveza de grifo": "Cerveza de grifo", "Vino": "Vino", "Botella y lata": "Botella y lata",
             "Bebidas": "Bebidas sin alcohol", "Café & té": "Café", "Focaccias": "Comida", "Pizzas": "Comida",
             "Tablas y más": "Comida"}
        return m.get(cat, cat or "Otros")

    grupos = defaultdict(float)
    comida = defaultdict(float)
    cervezas = defaultdict(lambda: [0.0, 0.0])
    pintas = canas = otras = 0.0
    prod_total = 0.0
    for r in D.get("prod", []):
        q, g = float(r.get("q") or 0), float(r.get("gross") or 0)
        if q <= 0 and abs(g) < 0.004:
            continue
        item = r.get("item") or ""
        cat = r.get("cat") or ""
        gr = grupo(cat, item)
        grupos[gr] += g
        prod_total += g
        if gr == "Comida":
            comida[item] += q
        if gr == "Cerveza de grifo":
            c = cervezas[limpia_cerveza(item)]
            c[0] += q
            c[1] += g
            v_ = (r.get("var") or "").lower()
            if v_ == "pinta":
                pintas += q
            elif v_ in ("caña", "cana"):
                canas += q
            else:
                otras += q
    if D.get("prod") and abs(prod_total - T["gross"]) > TOL:
        flag("error", f"El detalle de productos ({fmt(prod_total)} €) no cuadra con las ventas brutas ({fmt(T['gross'])} €).")

    # --- horas / concierto
    conc = {}
    for d in activos:
        if d in sup and d in horas:
            h0 = min(sup[d])
            tk_v = sum(x["tk"] for h, x in horas[d].items() if h >= h0)
            base_v = sum(x["gross"] - x["ret"] for h, x in horas[d].items() if h >= h0)
            h_ult = max((h for h, x in horas[d].items() if x["tk"] > 0), default=h0)
            q = sum(x[0] for x in sup[d].values())
            imp_base = sum(x[1] for x in sup[d].values())
            conc[d] = dict(h0=h0, h1=h_ult + 1, tk=int(round(tk_v)), base=round(base_v, 2), q=int(round(q)),
                           imp_base=round(imp_base, 2),
                           por_hora=[(h, int(round(x[0]))) for h, x in sorted(sup[d].items())])
    if horas:
        gh = sum(x["gross"] for d in horas for x in horas[d].values())
        if abs(gh - T["gross"]) > TOL:
            flag("error", f"El detalle por horas ({fmt(gh)} €) no cuadra con las ventas brutas ({fmt(T['gross'])} €).")

    # --- avisos a revisar
    for d in activos:
        for h, x in sorted(horas.get(d, {}).items()):
            if x["ret"] > 0.004:
                flag("info", f"Devolución el {dia_nombre(d)} {d[8:]}/{d[5:7]} a las {h} h: {fmt(x['ret'] * mult)} € con IVA "
                             f"({fmt(x['ret'])} € de base). Ya está descontada. Si es de una venta de otra semana, hay que ajustarla a mano.")
        if not horas.get(d) and V[d]["ret"] > 0.004:
            flag("info", f"Devoluciones el {dia_nombre(d)}: {fmt(V[d]['ret'] * mult)} € con IVA. Ya están descontadas.")
        for s, x in sorted(por_src[d].items()):
            if abs(x["ref"]) > 0.004:
                orig = f"«{s}»" if s else "sin origen"
                flag("info", f"Reembolsos por importe el {dia_nombre(d)}: {fmt(abs(x['ref']))} € en {int(round(x['n']))} pedidos de origen {orig}. "
                             f"No son ventas de la semana y no restan de la facturación; suelen ser pruebas o correcciones.")
            elif s not in ("", "Point of Sale", "Pedidos QR") and x["n"] > 0:
                flag("info", f"El {dia_nombre(d)} hay {int(round(x['n']))} pedidos de origen «{s}», distinto del TPV y del pedido QR.")
    qr_tk = sum(por_src[d].get("Pedidos QR", {}).get("tk", 0) for d in activos)
    qr_total = sum(por_src[d].get("Pedidos QR", {}).get("net", 0) + por_src[d].get("Pedidos QR", {}).get("tax", 0)
                   for d in activos)

    return dict(activos=activos, V=V, T=T, rate=rate, rate_txt=rate_txt, mult=mult, grupos=dict(grupos),
                prod_total=prod_total, comida=dict(comida), cervezas={k: tuple(v) for k, v in cervezas.items()},
                pintas=int(round(pintas)), canas=int(round(canas)), otras=int(round(otras)), horas=horas,
                conc=conc, flags=flags, qr_tk=int(round(qr_tk)), qr_total=round(qr_total, 2))


# ---------------------------------------------------------------- HTML
def kpi(label, valor, nota, main=False):
    return (f'<div class="kpi{" main" if main else ""}"><span class="l">{esc(label)}</span>'
            f'<span class="v">{esc(valor)}</span><span class="n">{esc(nota)}</span></div>')


def fila(nombre, ancho, valor, naranja=False, tenue=False):
    if ancho is None:
        return (f'<div class="row" role="listitem"><span class="name">{esc(nombre)}</span><span class="track"></span>'
                f'<span class="val dim">{esc(valor)}</span></div>')
    return (f'<div class="row" role="listitem"><span class="name">{esc(nombre)}</span><span class="track">'
            f'<span class="fill{" o" if naranja else ""}" style="width:{ancho:.1f}%"></span></span>'
            f'<span class="val{" dim" if tenue else ""}">{esc(valor)}</span></div>')


def tabla(cabeceras, filas, total=False, extra=None):
    h = "<thead><tr>" + "".join(f"<th>{esc(c)}</th>" for c in cabeceras) + "</tr></thead>"
    b = []
    for i, f in enumerate(filas):
        cls = ' class="total"' if total and i == len(filas) - 1 else ""
        celdas = []
        for j, c in enumerate(f):
            cls_td = ' class="dim"' if (isinstance(c, str) and c == "—") else ""
            celdas.append(f"<td{cls_td}>{c if (extra and j in extra) else esc(c)}</td>")
        b.append(f"<tr{cls}>" + "".join(celdas) + "</tr>")
    return f'<div class="card scroll"><table>{h}<tbody>{"".join(b)}</tbody></table></div>'


def construir_html(C, D, fuentes_css, generado):
    A, V, T = C["activos"], C["V"], C["T"]
    periodo = etiqueta_periodo(A)
    ini, fin = D.get("ini"), D.get("fin")
    rt = C["rate_txt"]
    mult = C["mult"]
    secs = []

    # Resumen
    bruto_iva = T["gross"] * mult
    tm = bruto_iva / T["tk"] if T["tk"] else 0
    k = [kpi("Facturación", fmt(T["total"]) + " €", "IVA incluido", main=True),
         kpi("Base imponible", fmt(T["base"]) + " €", f"IVA {rt} %: {fmt(T['iva'])} €"),
         kpi("Tickets", str(T["tk"]), f"Ticket medio {fmt(tm)} €")]
    if T["total"] > 0:
        k.append(kpi("Efectivo", fmt(T["efectivo"]) + " €", f"{pct(100 * T['efectivo'] / T['total'], 0)} del total"))
        k.append(kpi("Datáfono", fmt(T["datafono"]) + " €", f"{pct(100 * T['datafono'] / T['total'], 0)} del total"))
    vasos = C["pintas"] + C["canas"] + C["otras"]
    if vasos:
        nota = f"{C['pintas']} pintas y {C['canas']} cañas" + (f" y {C['otras']} otros" if C["otras"] else "")
        k.append(kpi("Cervezas servidas", str(vasos), nota))
    if C["conc"]:
        qtot = sum(c["q"] for c in C["conc"].values())
        dias_c = lista_y(dia_nombre(d).capitalize() for d in C["conc"])
        imp = sum(c["imp_base"] for c in C["conc"].values()) * mult
        pp = imp / qtot if qtot else 0
        k.append(kpi("Suplementos de concierto", str(qtot), f"{dias_c}, {fmt(pp)} € por persona"))
    secs.append('<section id="resumen"><h2>Resumen</h2><div class="kpis">' + "".join(k) + "</div></section>")

    # Facturación por día
    filas = [[dia_largo(d), V[d]["tk"], fmt(V[d]["base"]), fmt(V[d]["iva"]), fmt(V[d]["total"]),
              fmt(V[d]["efectivo"]), fmt(V[d]["datafono"])] for d in A]
    filas.append(["Total", T["tk"], fmt(T["base"]), fmt(T["iva"]), fmt(T["total"]), fmt(T["efectivo"]), fmt(T["datafono"])])
    devs = []
    for d in A:
        for h, x in sorted(C["horas"].get(d, {}).items()):
            if x["ret"] > 0.004:
                devs.append(f"{dia_nombre(d)} a las {h} h ({fmt(x['ret'] * mult)} €)")
    nota_dev = (f'<p class="small">Incluye devoluciones ya descontadas: {esc(lista_y(devs))}.</p>' if devs else "")
    secs.append('<section id="facturacion"><h2>Facturación por día</h2>'
                '<p class="lead">Las cifras que van a Odoo: ventas del día menos devoluciones. El datáfono es el resto hasta el total.</p>'
                + tabla(["Día", "Tickets", "Base", f"IVA {rt} %", "Total", "Efectivo", "Datáfono"], filas, total=True)
                + nota_dev + "</section>")

    # Hora a hora
    for d in A:
        hs = C["horas"].get(d)
        if not hs or V[d]["tk"] < UMBRAL_HORARIO:
            continue
        con_ventas = [h for h, x in hs.items() if x["tk"] > 0]
        if not con_ventas:
            continue
        h_min, h_max = min(con_ventas), max(con_ventas)
        neto = {h: hs.get(h, {"gross": 0, "ret": 0, "tk": 0}) for h in range(h_min, h_max + 1)}
        netos = {h: x["gross"] - x["ret"] for h, x in neto.items()}
        mx = max(netos.values()) or 1
        pico = max(netos, key=lambda h: netos[h])
        cc = C["conc"].get(d)
        filas_h = []
        for h in range(h_min, h_max + 1):
            x = neto[h]
            if x["tk"] <= 0 and abs(netos[h]) < 0.004:
                filas_h.append(fila(f"{h} h", None, "sin ventas"))
            else:
                filas_h.append(fila(f"{h} h", max(0, netos[h]) / mx * 100, f"{fmt(netos[h])} · {int(round(x['tk']))} tk",
                                    naranja=bool(cc and h >= cc["h0"])))
        if cc and V[d]["base"] > 0:
            share = 100 * cc["base"] / V[d]["base"]
            lead = (f"El {pct(share, 0)} de lo vendido el {dia_nombre(d)} cayó entre las {cc['h0']} y las {cc['h1']} h, "
                    f"con el pico a las {pico} h. Importe sin IVA y tickets por hora.")
            pie = (f"En naranja, la franja del concierto: {cc['tk']} tickets y {fmt(cc['base'])} € de base. "
                   f"Antes de las {h_min} h no hubo ventas. El importe de cada hora es neto de devoluciones.")
        else:
            lead = f"El pico fue a las {pico} h. Importe sin IVA y tickets por hora."
            pie = f"Antes de las {h_min} h no hubo ventas. El importe de cada hora es neto de devoluciones."
        secs.append(f'<section id="horas-{d}"><h2>El {esc(dia_nombre(d))}, hora a hora</h2><p class="lead">{esc(lead)}</p>'
                    f'<div class="card bars" role="list" aria-label="Ventas por hora">{"".join(filas_h)}</div>'
                    f'<p class="small">{esc(pie)}</p></section>')

    # Concierto y afluencia
    if C["conc"]:
        qtot = sum(c["q"] for c in C["conc"].values())
        tktot = sum(c["tk"] for c in C["conc"].values())
        imp = sum(c["imp_base"] for c in C["conc"].values())
        det = []
        for d, c in C["conc"].items():
            det.append(" y ".join(f"{q} a las {h} h" for h, q in c["por_hora"]))
        rango = lista_y(f"de {c['h0']} a {c['h1']} h" for c in C["conc"].values())
        kp = [kpi("Personas con suplemento", str(qtot), "; ".join(det)),
              kpi("Importe del suplemento", fmt(imp * mult) + " €", f"{fmt(imp)} € de base"),
              kpi("Tickets en el concierto", str(tktot), rango.capitalize())]
        if tktot > qtot:
            aviso = (f"Hubo {tktot} tickets durante el concierto y solo {qtot} suplementos. Si no todo el público lo pagó, "
                     f"la afluencia real es mayor que {qtot}. Para saber el aforo hay que compararlo con el recuento de la puerta.")
        else:
            aviso = (f"Hubo {qtot} suplementos y {tktot} tickets durante el concierto. Conviene compararlo con el recuento de la puerta.")
        secs.append('<section id="concierto"><h2>Concierto y afluencia</h2>'
                    '<p class="lead">El suplemento de concierto hace el papel de los tokens: cada unidad es una persona que pagó para entrar.</p>'
                    f'<div class="kpis">{"".join(kp)}</div>'
                    f'<div class="note"><p><span class="pill warn">A contrastar</span></p><p>{esc(aviso)}</p></div></section>')

    # Cuadre
    cu = []
    for d in A:
        if V[d]["efectivo"] > 0.004:
            cu.append([f"Efectivo del {dia_nombre(d)}", fmt(V[d]["efectivo"]), "—"])
        if V[d]["datafono"] > 0.004:
            cu.append([f"Datáfono del {dia_nombre(d)}", fmt(V[d]["datafono"]), "—"])
    det_tk = ", ".join(f"{V[d]['tk']} {dia_nombre(d)}" for d in A)
    cu.append([f"Tickets ({det_tk})", str(T["tk"]), "—"])
    if T["ret"] > 0.004:
        cu.append(["Devoluciones (con IVA)", fmt(T["ret"] * mult), "—"])
    if C["conc"]:
        cu.append(["Suplementos de concierto", str(sum(c["q"] for c in C["conc"].values())), "—"])
    secs.append('<section id="cuadre"><h2>Cuadre con lo anotado</h2>'
                '<p class="lead">Estas son las cifras de Square para comparar una a una con lo que se apunta en el local.</p>'
                + tabla(["Concepto", "Square", "Anotado"], cu) + "</section>")

    # Square vs facturación (solo si hay reembolsos por importe)
    if abs(T["ref"]) > 0.004:
        filas_r = [["Cobrado según el panel de Square", fmt(T["cobrado"])],
                   ["Se suman los reembolsos por importe, que no son ventas de la semana", "+" + fmt(-T["ref"])],
                   ["Facturación", fmt(T["total"])]]
        secs.append('<section id="diferencia"><h2>Por qué Square puede enseñar otra cifra</h2>'
                    f'<p class="lead">El panel de Square enseña {esc(fmt(T["cobrado"]))} € cobrados. Es menor que la facturación porque se restan '
                    'reembolsos por importe que no corresponden a ventas de esta semana. Así se llega de una cifra a otra.</p>'
                    + tabla(["Concepto", "Importe"], filas_r, total=True) + "</section>")

    # Producto
    if C["grupos"] and C["prod_total"] > 0:
        orden = sorted(C["grupos"].items(), key=lambda kv: -kv[1])
        mx = orden[0][1] or 1
        filas_p = [fila(n, g / mx * 100, f"{fmt(g)} · {pct(100 * g / C['prod_total'])}", naranja=(n == "Suplemento concierto"))
                   for n, g in orden if g > 0.004]
        nota_c = ""
        if C["comida"]:
            nota_c = ("Comida: " + lista_y(f"{int(q)} {n}" for n, q in sorted(C["comida"].items(), key=lambda kv: -kv[1])) + ".")
        secs.append('<section id="producto"><h2>Qué se vendió</h2>'
                    f'<p class="lead">Importe sin IVA por tipo de producto, antes de devoluciones: {esc(fmt(C["prod_total"]))} €.</p>'
                    f'<div class="card bars" role="list" aria-label="Ventas por tipo de producto">{"".join(filas_p)}</div>'
                    + (f'<p class="small">{esc(nota_c)}</p>' if nota_c else "") + "</section>")

    # Cervezas
    if C["cervezas"]:
        orden = sorted(C["cervezas"].items(), key=lambda kv: (-kv[1][0], -kv[1][1]))
        mx = orden[0][1][0] or 1
        vasos = sum(v[0] for _, v in orden)
        top2 = sum(v[0] for _, v in orden[:2])
        lead = f"{int(vasos)} vasos entre pintas y cañas."
        if len(orden) >= 2:
            lead += f" {orden[0][0]} y {orden[1][0]} suman el {pct(100 * top2 / vasos, 0)}."
        filas_c = [fila(n, v[0] / mx * 100, f"{int(v[0])} · {fmt(v[1])} €") for n, v in orden]
        secs.append('<section id="cervezas"><h2>Cervezas de grifo</h2>'
                    f'<p class="lead">{esc(lead)}</p><div class="card bars" role="list" aria-label="Vasos de cerveza por variedad">'
                    f'{"".join(filas_c)}</div><p class="small">Vasos y euros sin IVA. No incluye la tabla de degustación.</p></section>')

    # Lectura rápida
    pts = []
    pri = max(A, key=lambda d: V[d]["total"])
    if C["conc"]:
        d0 = max(C["conc"], key=lambda d: C["conc"][d]["base"])
        c0 = C["conc"][d0]
        pts.append(f"El negocio de la semana es el concierto del {dia_nombre(d0)} por la noche: {fmt(c0['base'])} € de base en {c0['h1'] - c0['h0']} horas.")
    elif T["total"] > 0:
        pts.append(f"El día más fuerte fue el {dia_nombre(pri)}: {fmt(V[pri]['total'])} € con IVA, el {pct(100 * V[pri]['total'] / T['total'], 0)} del total.")
    for d in A:
        if d != (d0 if C["conc"] else pri) and V[d]["tk"] > 0:
            pts.append(f"El {dia_nombre(d)} tuvo {V[d]['tk']} tickets y {fmt(V[d]['total'])} €.")
    if C["qr_tk"] > 0 and T["total"] > 0:
        palabra = "Solo " if C["qr_tk"] < T["tk"] / 2 else ""
        pts.append(f"{palabra}{C['qr_tk']} de los {T['tk']} tickets entraron por el pedido QR ({fmt(C['qr_total'])} €, el "
                   f"{pct(100 * C['qr_total'] / T['total'], 0)} de la facturación). El resto se cobró en el TPV.")
    if C["grupos"].get("Comida") and C["prod_total"] > 0:
        pts.append(f"La comida es el {pct(100 * C['grupos']['Comida'] / C['prod_total'], 0)} de lo vendido.")
    if T["tk"]:
        pts.append(f"El ticket medio es de {fmt(tm)} € con IVA.")
    secs.append('<section id="lectura"><h2>Lectura rápida</h2><ul class="plain">'
                + "".join(f"<li>{esc(p)}</li>" for p in pts) + "</ul></section>")

    # A revisar
    if C["flags"]:
        hay_error = any(f["nivel"] == "error" for f in C["flags"])
        pill = '<span class="pill bad">Error de datos</span>' if hay_error else '<span class="pill warn">A revisar</span>'
        secs.append('<section id="revisar"><h2>Puntos a revisar</h2>'
                    f'<div class="note{" bad" if hay_error else ""}"><p>{pill}</p>'
                    + "".join(f"<p>{esc(f['texto'])}</p>" for f in C["flags"]) + "</div></section>")
    else:
        secs.append('<section id="revisar"><h2>Puntos a revisar</h2><div class="note"><p><span class="pill ok">Todo cuadra</span></p>'
                    '<p>Los cobros de Square cuadran con las ventas y no hay devoluciones ni reembolsos.</p></div></section>')

    pie = ("<strong>Cómo se han calculado las cifras.</strong> Fuente: Square, ubicación ALTHAIA ARTESANA SL (Altea), horario de Madrid; "
           f"el día contable empieza a las 00:00. Se cuentan los pedidos cerrados con IVA incluido en los precios. Facturación = ventas del periodo "
           f"menos devoluciones; los reembolsos por importe no restan. El ticket medio usa las ventas antes de devoluciones "
           f"({fmt(bruto_iva)} € entre {T['tk']} tickets). Los datos de Square se actualizan cada 15 minutos aproximadamente.")
    cabecera = (f'<header><span class="eyebrow">Taproom Althaia × Pit</span><h1>{esc(periodo)}</h1>'
                f'<p class="sub">Caja de ALTHAIA ARTESANA SL en Square. Datos extraídos el {esc(generado)}. Importes en euros.</p></header>')
    return ('<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">'
            f'<title>Informe taproom Althaia</title><style>{fuentes_css}{CSS}</style></head><body><div class="wrap">'
            + cabecera + "".join(secs) + f'<footer><p class="small">{pie}</p></footer></div></body></html>')


# ---------------------------------------------------------------- fuentes y PDF
def css_fuentes():
    base = "/tmp/fonts-npm/node_modules/@fontsource"
    test = f"{base}/anton/files/anton-latin-400-normal.woff2"
    if not os.path.exists(test):
        try:
            subprocess.run(["npm", "install", "--prefix", "/tmp/fonts-npm", "--silent", "@fontsource/anton",
                            "@fontsource/signika", "@fontsource/space-mono"], timeout=180, check=True,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception:
            pass
    if not os.path.exists(test):
        return "@page { size:160mm 285mm; margin:0; }"
    faces = [("Anton", 400, "anton", "anton-latin-400-normal"), ("Signika", 400, "signika", "signika-latin-400-normal"),
             ("Signika", 600, "signika", "signika-latin-600-normal"), ("Space Mono", 400, "space-mono", "space-mono-latin-400-normal"),
             ("Space Mono", 700, "space-mono", "space-mono-latin-700-normal")]
    css = "".join(f"@font-face {{ font-family:'{n}'; font-weight:{w}; src:url('file://{base}/{p}/files/{f}.woff2') format('woff2'); }}\n"
                  for n, w, p, f in faces)
    return css + "@page { size:160mm 285mm; margin:0; }\n"


def a_pdf(html_path, pdf_path):
    chrome = shutil.which("google-chrome") or shutil.which("chromium") or shutil.which("chromium-browser")
    if not chrome:
        raise RuntimeError("No hay Chrome/Chromium para generar el PDF")
    subprocess.run([chrome, "--headless=new", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer",
                    "--run-all-compositor-stages-before-draw", "--virtual-time-budget=8000",
                    f"--print-to-pdf={pdf_path}", f"file://{html_path}"], check=True, timeout=120,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


# ---------------------------------------------------------------- Excel / CSV
def asiento(C):
    """Filas de importacion de Odoo (un asiento por dia). Devuelve (filas, debe, haber)."""
    filas, debe, haber = [], 0.0, 0.0
    A, V, rt = C["activos"], C["V"], C["rate_txt"]

    def dc(x, naturaleza):
        if naturaleza == "debe":
            return (round(x, 2), 0.0) if x >= 0 else (0.0, round(-x, 2))
        return (0.0, round(x, 2)) if x >= 0 else (round(-x, 2), 0.0)
    for d in A:
        v = V[d]
        lineas = []
        if abs(v["efectivo"]) > 0.004:
            lineas.append((CUENTAS["efectivo"], "Cobros en efectivo") + dc(v["efectivo"], "debe"))
        if abs(v["datafono"]) > 0.004:
            lineas.append((CUENTAS["datafono"], "Cobros con datáfono") + dc(v["datafono"], "debe"))
        lineas.append((CUENTAS["ventas"], "Ventas (base imponible)") + dc(v["base"], "haber"))
        lineas.append((CUENTAS["iva"], f"IVA repercutido {rt} %") + dc(v["iva"], "haber"))
        for i, (cta, etq, de, ha) in enumerate(lineas):
            if i == 0:
                filas.append([d, CUENTAS["diario"], f"Ventas taproom Althaia {fecha_corta(d)}", cta, etq, de, ha])
            else:
                filas.append(["", "", "", cta, etq, de, ha])
            debe += de
            haber += ha
    return filas, round(debe, 2), round(haber, 2)


def crear_excel(C, D, ruta, filas_odoo):
    from openpyxl import Workbook
    from openpyxl.comments import Comment
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    F = "Arial"
    hdr_font = Font(name=F, bold=True, color="FFFFFF", size=10)
    hdr_fill = PatternFill("solid", fgColor="12234A")
    azul = Font(name=F, color="0000FF", size=10)
    negro = Font(name=F, size=10)
    negrita = Font(name=F, size=10, bold=True)
    gris = PatternFill("solid", fgColor="F2F2F2")
    lado = Side(style="thin", color="BFBFBF")
    caja = Border(top=lado, bottom=lado, left=lado, right=lado)
    EUR = '#,##0.00;(#,##0.00);"-"'
    A, V, T = C["activos"], C["V"], C["T"]

    wb = Workbook()
    imp = wb.active
    imp.title = "Importar Odoo"
    res = wb.create_sheet("Resumen")
    nts = wb.create_sheet("Notas")

    # Importar Odoo (valores fijos: Odoo lee la primera hoja)
    cab = ["date", "journal_id", "ref", "line_ids/account_id", "line_ids/name", "line_ids/debit", "line_ids/credit"]
    for j, h in enumerate(cab, 1):
        c = imp.cell(1, j, h)
        c.font, c.fill = hdr_font, hdr_fill
    for i, f in enumerate(filas_odoo, 2):
        for j, v in enumerate(f, 1):
            c = imp.cell(i, j, v if v != "" else None)
            c.font = negro
            if j in (6, 7):
                c.number_format = "0.00"
            if j in (4,):
                c.number_format = "@"
    for col, w in zip("ABCDEFG", (12, 20, 36, 20, 28, 14, 14)):
        imp.column_dimensions[col].width = w
    imp.freeze_panes = "A2"
    ult_imp = len(filas_odoo) + 1

    # Resumen
    res["A1"] = f"Taproom Althaia — {etiqueta_periodo(A).lower()} de {dt.date.fromisoformat(A[-1]).year}"
    res["A1"].font = Font(name=F, bold=True, size=13, color="12234A")
    res["A2"] = "Fuente: Square (cuenta Althaia). Importes con IVA incluido. Azul = dato de entrada, negro = fórmula."
    res["A2"].font = Font(name=F, italic=True, size=9, color="595959")
    heads = ["Día", "Fecha", "Tickets", "Base", f"IVA {C['rate_txt']} %", "Total", "Efectivo", "Datáfono", "Cuadre (debe ser 0)"]
    HR = 4
    for j, h in enumerate(heads, 1):
        c = res.cell(HR, j, h)
        c.font, c.fill, c.border = hdr_font, hdr_fill, caja
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    res.row_dimensions[HR].height = 30
    for i, d in enumerate(A):
        r = HR + 1 + i
        v = V[d]
        res.cell(r, 1, dia_largo(d)).font = negro
        c = res.cell(r, 2, d)
        c.font, c.number_format = azul, "@"
        c = res.cell(r, 3, v["tk"])
        c.font, c.number_format = azul, "0"
        for col, val in ((4, v["base"]), (5, v["iva"]), (7, v["efectivo"]), (8, v["datafono"])):
            c = res.cell(r, col, val)
            c.font, c.number_format = azul, EUR
        c = res.cell(r, 6, f"=D{r}+E{r}")
        c.font, c.number_format = negro, EUR
        c = res.cell(r, 9, f"=ROUND(G{r}+H{r}-F{r},2)")
        c.font, c.number_format = negro, EUR
        for col in range(1, 10):
            res.cell(r, col).border = caja
    n = len(A)
    TR = HR + 1 + n
    res.cell(TR, 1, "Total").font = negrita
    for col in range(3, 10):
        L = "ABCDEFGHI"[col - 1]
        c = res.cell(TR, col, f"=SUM({L}{HR + 1}:{L}{TR - 1})")
        c.font, c.number_format = negrita, ("0" if col == 3 else EUR)
    for col in range(1, 10):
        res.cell(TR, col).border, res.cell(TR, col).fill = caja, gris
    r0 = TR + 2
    res.cell(r0, 1, "Comprobaciones").font = negrita
    res.cell(r0 + 1, 1, "Total facturación del informe (con IVA)").font = negro
    c = res.cell(r0 + 1, 2, T["total"])
    c.font, c.number_format = azul, EUR
    res.cell(r0 + 2, 1, "Diferencia con el total de este Excel").font = negro
    c = res.cell(r0 + 2, 2, f"=ROUND(F{TR}-B{r0 + 1},2)")
    c.number_format = EUR
    res.cell(r0 + 3, 1, "Asientos: suma Debe − suma Haber").font = negro
    c = res.cell(r0 + 3, 2, f"=ROUND(SUM('Importar Odoo'!F2:F{ult_imp})-SUM('Importar Odoo'!G2:G{ult_imp}),2)")
    c.number_format = EUR
    res.cell(r0 + 4, 1, "Estado").font = negrita
    c = res.cell(r0 + 4, 2, f'=IF(AND(B{r0 + 2}=0,B{r0 + 3}=0,I{TR}=0),"OK","REVISAR")')
    c.font = negrita
    for col, w in zip("ABCDEFGHI", (40, 22, 9, 12, 12, 12, 12, 12, 16)):
        res.column_dimensions[col].width = w
    res.sheet_view.showGridLines = False
    res.cell(r0 + 4, 3).value = None

    # Notas
    notas = [("Cómo importarlo en Odoo", True),
             ("1. Contabilidad → Asientos contables → botón de favoritos (⭐) → Importar registros.", False),
             ("2. Subir este archivo (solo se lee la primera hoja, 'Importar Odoo') o el CSV equivalente.", False),
             ("3. Comprobar que las 7 columnas se asignan solas. Pulsar 'Probar' y, si no da errores, 'Importar'.", False),
             ("4. Los asientos quedan en borrador: revisar y publicar.", False), ("", False),
             ("Qué se ha asumido (confirmar con el gestor antes de publicar)", True),
             (f"• Diario '{CUENTAS['diario']}' y cuentas {CUENTAS['efectivo']} (efectivo), {CUENTAS['datafono']} (datáfono), "
              f"{CUENTAS['ventas']} (ventas) y {CUENTAS['iva']} (IVA repercutido). Si usáis otras, buscar y reemplazar en la columna D de la primera hoja.", False),
             ("• Un asiento por día con el resumen de ventas (base + IVA) y el cobro por efectivo y datáfono. No es ticket a ticket.", False),
             ("• Al ser asiento manual, Odoo no etiqueta el IVA para el modelo 303. Si lo necesitáis, hay que importar como factura de cliente con impuesto.", False),
             ("• El datáfono es el resto hasta el total del día. Si el gestor prefiere una cuenta puente hasta la liquidación bancaria, cambiar la cuenta en la columna D.", False),
             ("• Facturación = ventas menos devoluciones. Los reembolsos por importe no restan.", False), ("", False),
             ("Puntos a revisar esta semana", True)]
    if C["flags"]:
        notas += [("• " + f["texto"], False) for f in C["flags"]]
    else:
        notas.append(("• Nada: los cobros de Square cuadran con las ventas y no hay devoluciones ni reembolsos.", False))
    for i, (t, b) in enumerate(notas, 1):
        c = nts.cell(i, 1, t)
        c.font = Font(name=F, size=10, bold=b, color="12234A" if b else "000000")
        c.alignment = Alignment(wrap_text=True, vertical="top")
    nts.column_dimensions["A"].width = 110
    wb.calculation.fullCalcOnLoad = True
    wb.save(ruta)


def recalcular(ruta):
    """Recalcula con LibreOffice para que las fórmulas lleven valores en caché (si no se puede, se deja como está)."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return False
    tmp = os.path.join(os.path.dirname(ruta), "_recalc")
    os.makedirs(tmp, exist_ok=True)
    try:
        subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", tmp, ruta], check=True,
                       timeout=120, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        out = os.path.join(tmp, os.path.basename(ruta))
        from openpyxl import load_workbook
        wb = load_workbook(out, data_only=True)
        if wb["Resumen"].max_row and any(c.value is not None for row in wb["Resumen"].iter_rows(min_row=4, max_row=8) for c in row[5:6]):
            shutil.move(out, ruta)
            return True
    except Exception:
        pass
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return False


def crear_csv(ruta, filas):
    cab = ["date", "journal_id", "ref", "line_ids/account_id", "line_ids/name", "line_ids/debit", "line_ids/credit"]
    with open(ruta, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(cab)
        for r in filas:
            w.writerow([r[0], r[1], r[2], r[3], r[4], f"{r[5]:.2f}", f"{r[6]:.2f}"])


# ---------------------------------------------------------------- email
def texto_email(C, nombres):
    A, V, T = C["activos"], C["V"], C["T"]
    periodo = etiqueta_periodo(A)
    L = [f"Hola Jorge:", "",
         f"Te paso el informe del taproom ({periodo.lower()}): el PDF para leerlo, el Excel para subirlo a Odoo y el mismo contenido en CSV por si el Excel da guerra.",
         "", f"Facturación: {fmt(T['total'])} € con IVA (base {fmt(T['base'])} €, IVA {fmt(T['iva'])} €).",
         f"Tickets: {T['tk']}. Efectivo {fmt(T['efectivo'])} €, datáfono {fmt(T['datafono'])} €."]
    if C["conc"]:
        q = sum(c["q"] for c in C["conc"].values())
        tkc = sum(c["tk"] for c in C["conc"].values())
        L.append(f"Concierto: {q} suplementos y {tkc} tickets en la franja del concierto (el aforo real hay que contrastarlo con la puerta).")
    L.append("")
    if C["flags"]:
        L.append("Cosas a mirar antes de subirlo a Odoo:")
        L += [f"- {f['texto']}" for f in C["flags"]]
    else:
        L.append("No hay nada raro: los cobros de Square cuadran con las ventas.")
    L += ["", "Para subirlo: Contabilidad > Asientos contables > favoritos > Importar registros. El diario y las cuentas "
              f"({CUENTAS['efectivo']}, {CUENTAS['datafono']}, {CUENTAS['ventas']}, {CUENTAS['iva']}) son supuestos: que los confirme el gestor.",
          "", "Cualquier duda me dices.", "", "Un abrazo,", "Pedro"]
    return "\n".join(L)


# ---------------------------------------------------------------- main
def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(2)
    D = json.load(open(sys.argv[1], encoding="utf-8"))
    out = os.path.abspath(sys.argv[2])
    os.makedirs(out, exist_ok=True)
    C = calcular(D)
    if C is None:
        print(json.dumps({"ok": True, "sin_ventas": True, "ini": D.get("ini"), "fin": D.get("fin")}, ensure_ascii=False))
        return
    A = C["activos"]
    nom = nombre_archivo_periodo(A)
    filas, debe, haber = asiento(C)
    if abs(debe - haber) > 0.005:
        C["flags"].append({"nivel": "error", "texto": f"Los asientos para Odoo no cuadran: debe {fmt(debe)} € y haber {fmt(haber)} €."})
    generado = dt.datetime.now().strftime("%d/%m/%Y a las %H:%M")
    html_path = os.path.join(out, "informe.html")
    open(html_path, "w", encoding="utf-8").write(construir_html(C, D, css_fuentes(), generado))
    pdf = os.path.join(out, f"Informe-Althaia-{nom}.pdf")
    a_pdf(html_path, pdf)
    xlsx = os.path.join(out, f"Ventas-Althaia-{nom}-Odoo.xlsx")
    crear_excel(C, D, xlsx, filas)
    cache = recalcular(xlsx)
    csv_path = os.path.join(out, f"Ventas-Althaia-{nom}-Odoo.csv")
    crear_csv(csv_path, filas)
    email = texto_email(C, None)
    open(os.path.join(out, "email.txt"), "w", encoding="utf-8").write(email)
    errores = [f["texto"] for f in C["flags"] if f["nivel"] == "error"]
    T = C["T"]
    meta = {"ok": not errores, "errores": errores, "avisos": [f["texto"] for f in C["flags"] if f["nivel"] == "info"],
            "periodo": etiqueta_periodo(A), "asunto": f"Informe taproom Althaia · {etiqueta_periodo(A)}",
            "total": T["total"], "base": T["base"], "iva": T["iva"], "tickets": T["tk"], "efectivo": T["efectivo"],
            "datafono": T["datafono"], "excel_con_valores_en_cache": cache,
            "archivos": [pdf, xlsx, csv_path], "email": os.path.join(out, "email.txt")}
    json.dump(meta, open(os.path.join(out, "meta.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    print(json.dumps(meta, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
