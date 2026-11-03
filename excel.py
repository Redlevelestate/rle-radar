#!/usr/bin/env python3
"""
RLE · Radar diario — candidatos, registro y Excel.

  python3 excel.py candidatos OUT_DIR registro.txt [max]   -> OUT_DIR/candidatos.json (a enriquecer por la IA)
  python3 excel.py construir  OUT_DIR registro.txt FECHA   -> OUT_DIR/Radar_RLE_FECHA.xlsx, registro_nuevo.txt, resumen.json

Ficheros de OUT_DIR (los que falten se toleran):
  borme.json, renovables.json, licitaciones.json       (radar.py, deterministas)
  web.json            {"noticias":[...], "empleo":[...]}   (búsqueda web de la IA)
  enriquecidas.json   {clave: {web,email,decisor,cargo,linkedin,telefono,motivo,fuente_email,en_crm}}
"""
import csv, json, os, re, sys
from collections import Counter
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from radar import norm_nombre, ES_EMPRESA


def cargar(p, defecto):
    try:
        return json.load(open(p))
    except Exception:
        return defecto


import hashlib


def H(clave):
    return hashlib.sha1(clave.encode("utf-8")).hexdigest()[:9]


class Registro:
    """Registro compacto: una línea por empresa vista -> '<hash9> <AAMMDD primera> <AAMMDD última> <E si enriquecida>'.
    Las enriquecidas se guardan siempre; el resto, 45 días (para no re-mostrarlas como nuevas)."""
    def __init__(self, p):
        self.d = {}
        if os.path.exists(p):
            for l in open(p, encoding="utf-8"):
                x = l.split()
                if len(x) >= 3:
                    self.d[x[0]] = {"p": x[1], "u": x[2], "e": len(x) > 3 and x[3] == "E"}

    def __contains__(self, clave):
        return bool(clave) and H(clave) in self.d

    def __len__(self):
        return len(self.d)

    def primera(self, clave):
        p = self.d[H(clave)]["p"]
        return f"{p[4:6]}/{p[2:4]}/20{p[0:2]}"

    def enriquecida(self, clave):
        return clave in self and self.d[H(clave)]["e"]

    def marcar(self, clave, fecha, enr=False):
        if not clave:
            return
        f = fecha.replace("-", "")[2:]
        k = H(clave)
        if k in self.d:
            self.d[k]["u"] = f; self.d[k]["e"] = self.d[k]["e"] or enr
        else:
            self.d[k] = {"p": f, "u": f, "e": enr}

    def guardar(self, p, fecha, dias=45):
        from datetime import date as _d, timedelta as _t
        lim = (_d.fromisoformat(fecha) - _t(days=dias)).strftime("%y%m%d")
        with open(p, "w", encoding="utf-8") as fh:
            for k, v in self.d.items():
                if v["e"] or v["u"] >= lim:
                    fh.write(f"{k} {v['p']} {v['u']}{' E' if v['e'] else ''}\n")


def leer_registro(p):
    return Registro(p)


def filas_fuentes(out):
    b = cargar(f"{out}/borme.json", {}).get("filas", [])
    r = cargar(f"{out}/renovables.json", {}).get("filas", [])
    l = cargar(f"{out}/licitaciones.json", {}).get("filas", [])
    w = cargar(f"{out}/web.json", {})
    return b, r, l, w.get("noticias", []), w.get("empleo", [])

# ------------------------------------------------------------------ candidatos a enriquecer

def candidatos(out, registro, maximo=45):
    reg = leer_registro(registro)
    b, r, l, n, e = filas_fuentes(out)
    cands = []

    def add(clave, empresa, prioridad, sector, señal, contexto, peso):
        if not clave or reg.enriquecida(clave):
            return
        cands.append({"clave": clave, "empresa": empresa, "prioridad": prioridad, "sector": sector,
                      "señal": señal, "contexto": contexto[:300], "peso": peso})

    for f in b:
        if f["prioridad"] != "A":
            continue
        # SPV nueva con socio único empresa -> se enriquece el GRUPO (es quien tiene web y decisores)
        if f.get("socio_es_empresa") and "Constitución" in f["señales"]:
            add(norm_nombre(f["socio_unico"]), f["socio_unico"], "A", f["sectores"],
                f"Crea nueva sociedad {f['empresa']} ({f['provincia']})", f["objeto_social"], 90)
            continue
        peso = 60 + min(30, int((f.get("ampliacion") or 0) / 500_000) * 5) + (10 if f.get("sector_claro") else 0)
        add(f["clave"], f["empresa"], "A", f["sectores"], f["señales"],
            f"{f['provincia']}. {f.get('objeto_social') or ''} Ampliación: {f.get('ampliacion') or ''}", peso)
    for f in r:
        if f.get("promotor"):
            add(f["clave"], f["promotor"], "A", "Energía y renovables", "Proyecto renovable en información pública (BOE)",
                f"{f['tecnologia']} {f['potencia']} {f['provincia']}", 80)
    for f in l:
        if f["prioridad"] == "A":
            add(f["clave"], f["empresa"], "A", "Suelo público", f["tipo"], f"{f['objeto']} — {f['organo']}", 95)
    for f in n:
        add(norm_nombre(f.get("empresa", "")), f.get("empresa", ""), "A", f.get("sector", ""), f.get("señal", "Noticia"),
            f.get("resumen", ""), 85)
    for f in e:
        add(norm_nombre(f.get("empresa", "")), f.get("empresa", ""), "A", f.get("sector", ""), "Busca perfil de expansión/suelo",
            f.get("puesto", ""), 88)
    vistos, out_c = set(), []
    for c in sorted(cands, key=lambda x: -x["peso"]):
        if c["clave"] not in vistos:
            vistos.add(c["clave"]); out_c.append(c)
    json.dump(out_c[:maximo], open(f"{out}/candidatos.json", "w"), ensure_ascii=False, indent=1)
    print(f"{len(out_c)} candidatos nuevos; se enriquecen {min(len(out_c), maximo)}")

# ------------------------------------------------------------------ Excel

def construir(out, registro, fecha):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import FormulaRule

    reg = leer_registro(registro)
    b, r, l, n, e = filas_fuentes(out)
    enr = cargar(f"{out}/enriquecidas.json", {})
    meta_b = cargar(f"{out}/borme.json", {})
    meta_l = cargar(f"{out}/licitaciones.json", {})
    alertas = cargar(f"{out}/alertas.json", [])

    ROJO, OSC, GRIS, CLARO = "8C1D2C", "1F2328", "6B7280", "F6F3F1"
    PRI = {"A": ("E7F4EA", "1E7B34"), "B": ("FFF4DB", "9A6700"), "C": ("F1F1F1", "6B7280")}
    thin = Side(style="thin", color="E5E1DE")
    wb = Workbook()

    def nueva(clave):
        return "Sí" if clave and clave not in reg else ("" if not clave else f"No (desde {reg.primera(clave)})")

    def hoja(titulo, cols, filas, anchos, pri_col=None, color=ROJO, nota=None):
        ws = wb.create_sheet(titulo)
        fila0 = 1
        if nota:
            ws.cell(1, 1, nota).font = Font(italic=True, color=GRIS, size=9)
            fila0 = 2
        for j, c in enumerate(cols, 1):
            cell = ws.cell(fila0, j, c)
            cell.font = Font(bold=True, color="FFFFFF", size=10)
            cell.fill = PatternFill("solid", fgColor=color)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        ws.row_dimensions[fila0].height = 30
        for i, f in enumerate(filas, fila0 + 1):
            for j, v in enumerate(f, 1):
                cell = ws.cell(i, j, v)
                cell.alignment = Alignment(vertical="top", wrap_text=isinstance(v, str) and len(v) > 40)
                cell.border = Border(bottom=thin)
                if isinstance(v, str) and re.match(r"^https?://", v):
                    cell.hyperlink = v; cell.font = Font(color="1A56DB", underline="single")
                elif isinstance(v, str) and v.startswith("www."):
                    cell.hyperlink = "https://" + v; cell.font = Font(color="1A56DB", underline="single")
                elif isinstance(v, (int, float)) and v and v >= 1000:
                    cell.number_format = '#,##0 "€"'
        for j, w in enumerate(anchos, 1):
            ws.column_dimensions[get_column_letter(j)].width = w
        ws.freeze_panes = ws.cell(fila0 + 1, 2)
        if filas:
            ws.auto_filter.ref = f"A{fila0}:{get_column_letter(len(cols))}{fila0 + len(filas)}"
            if pri_col:
                L = get_column_letter(pri_col)
                rng = f"A{fila0+1}:{get_column_letter(len(cols))}{fila0+len(filas)}"
                for p, (bg, fg) in PRI.items():
                    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'${L}{fila0+1}="{p}"'], fill=PatternFill("solid", fgColor=bg)))
        else:
            ws.cell(fila0 + 1, 1, "Sin resultados hoy en esta fuente.").font = Font(italic=True, color=GRIS)
        return ws

    # ---- ★ Para contactar (enriquecidas)
    contactar = []
    for clave, d in enr.items():
        if not (d.get("email") or d.get("decisor") or d.get("web")):
            continue
        contactar.append([
            d.get("empresa", ""), d.get("sector", ""), d.get("señal", ""), d.get("motivo", ""),
            d.get("web", ""), d.get("email", ""), d.get("fuente_email", ""), d.get("decisor", ""), d.get("cargo", ""),
            d.get("linkedin", ""), d.get("telefono", ""), d.get("provincia", ""), d.get("en_crm", ""), d.get("origen", ""),
        ])
    contactar.sort(key=lambda x: (x[5] == "", x[7] == ""))
    hojas_orden = []
    hoja("★ Para contactar",
         ["Empresa", "Sector", "Señal", "Por qué ahora", "Web", "Email", "Dónde se encontró el email", "Decisor", "Cargo",
          "LinkedIn", "Teléfono", "Provincia", "¿Ya en Zoho?", "Fuente"],
         contactar, [30, 18, 26, 46, 26, 30, 22, 24, 22, 30, 14, 14, 12, 16], color=ROJO,
         nota="Empresas nuevas del día con web/email/decisor encontrados. Emails sin validar todavía: pásalos por el validador antes de enviar.")

    # ---- BORME
    pri_ord = {"A": 0, "B": 1, "C": 2}
    bs = sorted(b, key=lambda f: (pri_ord[f["prioridad"]], -(f.get("ampliacion") or f.get("capital_constitucion") or 0)))
    hoja("BORME · señales",
         ["Prioridad", "Empresa", "Señales", "Sectores", "Sector claro", "Provincia", "Capital constitución", "Ampliación de capital",
          "Capital resultante", "Objeto social", "Domicilio", "Socio único", "Administradores", "Web (BORME)", "Nueva en el radar", "PDF BORME"],
         [[f["prioridad"], f["empresa"], f["señales"], f["sectores"], f.get("sector_claro", ""), f["provincia"],
           f.get("capital_constitucion"), f.get("ampliacion"), f.get("capital_resultante"), f.get("objeto_social", ""),
           f.get("domicilio", ""), f.get("socio_unico", ""), f.get("administradores", ""), f.get("web_borme", ""),
           nueva(f["clave"]), f["fuente_url"]] for f in bs],
         [9, 34, 26, 24, 8, 13, 13, 13, 13, 50, 34, 28, 40, 22, 14, 18], pri_col=1, color=OSC,
         nota="A = inversión fuerte o sociedad de un grupo · B = actividad relevante · C = señal débil o sociedad pequeña.")

    # ---- Carta física: constituciones del sector sin web
    carta = [f for f in bs if "Constitución" in f["señales"] and f.get("sector_claro") and not f.get("web_borme")]
    hoja("Carta física",
         ["Prioridad", "Empresa", "Sectores", "Provincia", "Domicilio social", "Administradores", "Socio único", "Capital", "Objeto social"],
         [[f["prioridad"], f["empresa"], f["sectores"], f["provincia"], f.get("domicilio", ""), f.get("administradores", ""),
           f.get("socio_unico", ""), f.get("capital_constitucion"), f.get("objeto_social", "")] for f in carta],
         [9, 34, 24, 13, 40, 40, 28, 12, 60], pri_col=1, color=OSC,
         nota="Sociedades recién creadas del sector: casi nunca tienen web ni email todavía. Su vía es la carta al domicilio social.")

    # ---- Renovables
    hoja("Renovables (BOE)",
         ["Promotor", "Otros promotores", "Tecnología", "Potencia", "Provincia", "Anuncio", "Nueva en el radar", "Enlace"],
         [[f["promotor"], f["otros_promotores"], f["tecnologia"], f["potencia"], f["provincia"], f["titulo"],
           nueva(f["clave"]), f["url"]] for f in r],
         [34, 30, 18, 12, 14, 80, 14, 20], color=OSC,
         nota="Proyectos en información pública: cada uno necesita suelo, servidumbres o subestación.")

    # ---- Licitaciones
    hoja("Licitaciones y suelo",
         ["Prioridad", "Tipo", "Adjudicataria", "NIF", "Objeto", "Órgano", "Lugar", "Importe licitación", "Importe adjudicación",
          "Nueva en el radar", "Enlace"],
         [[f["prioridad"], f["tipo"], f["empresa"], f["nif"], f["objeto"], f["organo"], f["lugar"], f.get("importe_licitacion"),
           f.get("importe_adjudicacion"), nueva(f["clave"]), f["url"]] for f in sorted(l, key=lambda x: x["prioridad"])],
         [9, 26, 34, 12, 60, 30, 16, 15, 15, 14, 20], pri_col=1, color=OSC,
         nota="A = compró suelo público (comprador de suelo confirmado) · B = ganó una obra grande de edificación.")

    # ---- Noticias e inversores
    hoja("Noticias e inversores",
         ["Empresa", "Sector", "Señal", "Resumen", "Zona", "Fecha noticia", "Nueva en el radar", "Enlace"],
         [[f.get("empresa", ""), f.get("sector", ""), f.get("señal", ""), f.get("resumen", ""), f.get("zona", ""),
           f.get("fecha", ""), nueva(norm_nombre(f.get("empresa", ""))), f.get("url", "")] for f in n],
         [30, 20, 24, 60, 16, 12, 14, 30], color=OSC)

    # ---- Empleo
    hoja("Empleo · expansión",
         ["Empresa", "Sector", "Puesto", "Ubicación", "Fecha", "Nueva en el radar", "Enlace"],
         [[f.get("empresa", ""), f.get("sector", ""), f.get("puesto", ""), f.get("ubicacion", ""), f.get("fecha", ""),
           nueva(norm_nombre(f.get("empresa", ""))), f.get("url", "")] for f in e],
         [30, 20, 40, 18, 12, 14, 30], color=OSC,
         nota="Empresas que contratan perfiles de suelo, expansión o inversión: tienen la necesidad ahora.")

    # ---- Resumen (primera hoja)
    ws = wb["Sheet"]; ws.title = "Resumen"; wb.move_sheet(ws, offset=-(len(wb.sheetnames) - 1))
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 38
    for c in "CDEF":
        ws.column_dimensions[c].width = 18
    ws["B2"] = "RADAR DIARIO DE EMPRESAS"
    ws["B2"].font = Font(bold=True, size=18, color=ROJO)
    ws["B3"] = f"Red Level Estate · {date.fromisoformat(fecha).strftime('%d/%m/%Y')}"
    ws["B3"].font = Font(size=11, color=GRIS)
    cb = Counter(f["prioridad"] for f in b)
    nuevas_total = len({k for k in [f["clave"] for f in b] + [f["clave"] for f in r] + [f["clave"] for f in l]
                        + [norm_nombre(x.get("empresa", "")) for x in n + e] if k and k not in reg})
    kpis = [("Empresas nuevas en el radar", nuevas_total), ("Listas para contactar", len(contactar)),
            ("Con email encontrado", sum(1 for x in contactar if x[5])), ("Para carta física", len(carta))]
    # KPIs en fichas: número grande arriba, etiqueta debajo
    for j, (t, v) in enumerate(kpis):
        col = 2 + j
        c = ws.cell(5, col, v); c.font = Font(bold=True, size=22, color=ROJO)
        c.alignment = Alignment(horizontal="left"); c.fill = PatternFill("solid", fgColor=CLARO)
        c = ws.cell(6, col, t); c.font = Font(size=9, color=GRIS); c.alignment = Alignment(wrap_text=True, vertical="top")
        c.fill = PatternFill("solid", fgColor=CLARO)
    ws.row_dimensions[5].height = 34; ws.row_dimensions[6].height = 28
    fila = 7
    fila += 1
    ws.cell(fila, 2, "POR FUENTE").font = Font(bold=True, size=10, color=GRIS); fila += 1
    hdr = ["Fuente", "Total", "Prioridad A", "Prioridad B", "Prioridad C"]
    for j, h in enumerate(hdr):
        c = ws.cell(fila, 2 + j, h); c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=OSC)
    fila += 1
    cl = Counter(f["prioridad"] for f in l)
    tabla = [
        (f"BORME ({meta_b.get('pdfs', 0)} provincias, {meta_b.get('actos_leidos', 0)} actos leídos)", len(b), cb["A"], cb["B"], cb["C"]),
        ("Renovables en información pública (BOE)", len(r), len(r), "", ""),
        ("Licitaciones y suelo público", len(l), cl["A"], cl["B"], ""),
        ("Noticias de expansión e inversores", len(n), len(n), "", ""),
        ("Ofertas de empleo de expansión/suelo", len(e), len(e), "", ""),
    ]
    for t in tabla:
        for j, v in enumerate(t):
            c = ws.cell(fila, 2 + j, v); c.border = Border(bottom=thin)
            if j: c.alignment = Alignment(horizontal="center")
        fila += 1
    if alertas:
        fila += 1
        ws.cell(fila, 2, "AVISOS").font = Font(bold=True, size=10, color=ROJO); fila += 1
        for a in alertas:
            ws.cell(fila, 2, "⚠ " + a).font = Font(color=ROJO); fila += 1
    fila += 1
    ws.cell(fila, 2, "CÓMO USARLO").font = Font(bold=True, size=10, color=GRIS); fila += 1
    for t in ["★ Para contactar: empresas nuevas con web, email y decisor. Es la hoja de trabajo.",
              "BORME · señales: todo lo detectado en el Registro Mercantil, ordenado por prioridad.",
              "Carta física: sociedades recién creadas del sector, sin web todavía.",
              "Renovables / Licitaciones / Noticias / Empleo: el resto de fuentes del día.",
              "«Nueva en el radar = No» indica que ya salió otro día (fecha entre paréntesis)."]:
        ws.cell(fila, 2, "· " + t).font = Font(size=10, color=OSC); fila += 1

    nombre = f"Radar_RLE_{fecha}.xlsx"
    wb.save(f"{out}/{nombre}")
    # Copia en CSV de cada hoja (texto: se sube a Drive como Google Sheet si el .xlsx no se puede subir)
    import csv as _csv
    os.makedirs(f"{out}/csv", exist_ok=True)
    for ws_ in wb.worksheets:
        if ws_.title == "Resumen":
            continue
        slug = re.sub(r"[^A-Za-z0-9]+", "_", ws_.title.replace("★", "")).strip("_")
        with open(f"{out}/csv/{slug}.csv", "w", newline="", encoding="utf-8") as fh:
            w = _csv.writer(fh)
            for row in ws_.iter_rows(values_only=True):
                if any(v not in (None, "") for v in row):
                    w.writerow(["" if v is None else v for v in row])

    # ---- registro actualizado
    for k in [f["clave"] for f in b] + [f["clave"] for f in r] + [f["clave"] for f in l] \
            + [norm_nombre(x.get("empresa", "")) for x in n + e]:
        reg.marcar(k, fecha)
    for k, d in enr.items():
        reg.marcar(k, fecha, enr=bool(d.get("web") or d.get("email")))
    reg.guardar(f"{out}/registro_nuevo.txt", fecha)

    resumen = {"fecha": fecha, "excel": nombre, "nuevas": nuevas_total, "contactar": len(contactar),
               "con_email": sum(1 for x in contactar if x[5]), "carta": len(carta),
               "borme": {"total": len(b), **cb, "pdfs": meta_b.get("pdfs", 0), "actos": meta_b.get("actos_leidos", 0)},
               "renovables": len(r), "licitaciones": len(l), "noticias": len(n), "empleo": len(e),
               "registro_total": len(reg), "alertas": alertas,
               "top": [{"empresa": x[0], "señal": x[2], "motivo": x[3], "web": x[4], "email": x[5], "decisor": x[7]} for x in contactar[:12]]}
    json.dump(resumen, open(f"{out}/resumen.json", "w"), ensure_ascii=False, indent=1)
    print(json.dumps({k: v for k, v in resumen.items() if k != "top"}, ensure_ascii=False))


if __name__ == "__main__":
    if sys.argv[1] == "candidatos":
        candidatos(sys.argv[2], sys.argv[3], int(sys.argv[4]) if len(sys.argv) > 4 else 45)
    elif sys.argv[1] == "construir":
        construir(sys.argv[2], sys.argv[3], sys.argv[4])
