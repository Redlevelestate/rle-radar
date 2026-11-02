#!/usr/bin/env python3
"""
RLE · Radar diario de empresas (Red Level Estate)

Fuentes oficiales deterministas (sin IA):
  borme   -> BORME Sección A: constituciones, ampliaciones de capital, nombramientos,
             cambios de objeto/domicilio, socio único, web de la sociedad.
  boe5b   -> BOE Sección V-B: información pública de instalaciones renovables
             (fotovoltaica, eólica, baterías, hidrógeno, biogás) con su promotor.
  placsp  -> Plataforma de Contratación: adjudicaciones de suelo público (enajenación
             de parcelas, derechos de superficie) y obras grandes de edificación.

Uso:
  python3 radar.py borme  AAAA-MM-DD  out/borme.json
  python3 radar.py boe5b  AAAA-MM-DD  out/renovables.json
  python3 radar.py placsp AAAA-MM-DD  out/licitaciones.json   (últimas 24 h hasta esa fecha)
  python3 radar.py parse-pdf fichero.pdf PROVINCIA            (prueba del parser)
"""
import json, os, re, subprocess, sys, tempfile, time, unicodedata
import urllib.request, urllib.error
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta, timezone

UA = "Mozilla/5.0 (radar diario de datos abiertos)"

# --------------------------------------------------------------------------- utilidades

def http_get(url, accept=None, retries=3, timeout=60):
    last = None
    for i in range(retries):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": UA, **({"Accept": accept} if accept else {})})
            with urllib.request.urlopen(req, timeout=timeout) as r:
                return r.read()
        except urllib.error.HTTPError as e:
            if e.code == 404:
                raise
            last = e
        except Exception as e:  # red, timeout
            last = e
        time.sleep(2 * (i + 1))
    raise last


def sin_acentos(s):
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def norm_nombre(s):
    s = sin_acentos(s or "").upper()
    s = re.sub(r"[\.,;:'\"()\-]", " ", s)
    s = re.sub(r"\b(SOCIEDAD LIMITADA( UNIPERSONAL| LABORAL| PROFESIONAL)?|SOCIEDAD ANONIMA( UNIPERSONAL)?|S L U|S A U|S L L|S L P|SLU|SAU|SLL|SLP|S L|S A|SL|SA|SCCL|S COOP|SOCIMI)\b", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def euros(txt):
    """'1.058.560,00 Euros' -> 1058560.0"""
    if not txt:
        return None
    m = re.search(r"([\d\.]+,\d{2})", txt)
    if not m:
        return None
    return float(m.group(1).replace(".", "").replace(",", "."))


def pdf_texto(data):
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
        f.write(data)
        p = f.name
    try:
        try:
            out = subprocess.run(["pdftotext", p, "-"], capture_output=True, timeout=120)
            if out.returncode == 0 and out.stdout.strip():
                return out.stdout.decode("utf-8", "replace")
        except FileNotFoundError:
            pass
        from pypdf import PdfReader  # respaldo
        return "\n".join((pg.extract_text() or "") for pg in PdfReader(p).pages)
    finally:
        os.unlink(p)


def walk(o):
    if isinstance(o, dict):
        yield o
        for v in o.values():
            yield from walk(v)
    elif isinstance(o, list):
        for v in o:
            yield from walk(v)


def as_text(v):
    if isinstance(v, dict):
        return v.get("texto") or v.get("#text") or ""
    return v or ""

# --------------------------------------------------------------------------- sectores

SECTORES = [
    ("Promoción inmobiliaria", r"promoci[oó]n inmobiliaria|promoci[oó]n de (edificaciones|viviendas|edificios|inmuebles)|promotora|desarrollo(s)? inmobiliario|construcci[oó]n de edificios|urbaniza|gesti[oó]n (urban[ií]stica|de suelo)|compra ?venta de (bienes )?inmuebles|adquisici[oó]n de (suelo|solares|terrenos|fincas)|\bsuelo\b|\bsolares\b|\bpromoci[oó]n(es)?\b|inmobiliari|bienes inmuebles|\binmuebles\b|rehabilitaci[oó]n de edificios|\bresidencial\b|\bbuild to rent\b|\bcoliving\b"),
    ("Industrial y logística", r"log[ií]stic|almacenaje|almacenamiento de mercanc|centro(s)? de distribuci|naves? industrial|parque(s)? empresarial|parque(s)? industrial|cross ?dock|[uú]ltima milla"),
    ("Energía y renovables", r"fotovolt|\bsolar\b|e[oó]lic|energ[ií]as? renovable|almacenamiento (de|el[eé]ctrico de) energ|bater[ií]as|hidr[oó]geno|biog[aá]s|biometano|producci[oó]n de energ[ií]a el[eé]ctrica|generaci[oó]n de energ|\bbess\b|\bhibridaci"),
    ("Centros de datos", r"centros? de datos|data ?cent|\bhyperscal|colocation"),
    ("Inversión inmobiliaria", r"\bsocimi\b|family office|inversi[oó]n inmobiliaria|fondo(s)? de inversi[oó]n|gesti[oó]n de patrimonios inmobiliarios|tenencia.{0,40}inmuebles|\bpatrimonial(es)?\b|real estate|\bcapital partners\b"),
    ("Hoteles y turismo", r"\bhotel|aparthotel|alojamientos? tur[ií]stic|\bhostel\b|camping"),
    ("Residencias y sanitario", r"residencia(s)? (de|para) (mayores|la tercera edad|ancianos|estudiantes)|residencia geri[aá]trica|centro(s)? socio ?sanitario|hospital|cl[ií]nica"),
    ("Retail y operadores", r"supermercado|hipermercado|gasolinera|estaci[oó]n(es)? de servicio|trasteros|self ?storage|gimnasio|centro(s)? deportivo|drive ?thru|lavado de veh[ií]culos|centro(s)? comercial"),
]
SECTORES_RE = [(n, re.compile(p, re.I)) for n, p in SECTORES]
# en el NOMBRE de la sociedad (más estricto, para actos sin objeto social)
NOMBRE_RE = [
    ("Promoción inmobiliaria", re.compile(r"PROMOC|PROMOTOR|INMOBILI|URBANI|DESARROLLOS|EDIFICACION|RESIDENCIAL|HOMES\b|LIVING\b|REAL ESTATE|PROPERT|SUELO|SOLARES|VIVIENDAS|BUILD", re.I)),
    ("Industrial y logística", re.compile(r"LOGIST|LOGIS\b|ALMACEN|NAVES|PARQUE EMPRESARIAL|PARQUE INDUSTRIAL|DISTRIBUCION", re.I)),
    ("Energía y renovables", re.compile(r"SOLAR|FOTOVOLT|\bFV\b|\bPV\b|EOLIC|RENOVABL|ENERG|WIND|POWER|STORAGE|BESS|HIDROGENO|HYDROGEN|BIOGAS|BIOMETANO", re.I)),
    ("Centros de datos", re.compile(r"DATA ?CENT|CENTRO DE DATOS|DATACENTER", re.I)),
    ("Inversión inmobiliaria", re.compile(r"SOCIMI|PATRIMON|INVERSIONES INMOB|CAPITAL PARTNERS|INVESTMENTS|ASSET", re.I)),
    ("Hoteles y turismo", re.compile(r"HOTEL|HOSPITALITY|APARTHOTEL|RESORT", re.I)),
    ("Residencias y sanitario", re.compile(r"RESIDENCIA|SENIOR|GERIATR|HOSPITAL|CLINIC", re.I)),
]


def sectores_de(objeto, nombre):
    """-> (sectores, fuerte). Fuerte = el sector está en el nombre o al principio del objeto social
    (los objetos genéricos que meten 'compraventa de inmuebles' entre veinte actividades son débiles)."""
    s, fuerte = [], False
    for n, rx in NOMBRE_RE:
        if nombre and rx.search(nombre):
            s.append(n); fuerte = True
    for n, rx in SECTORES_RE:
        if objeto and rx.search(objeto):
            if n not in s:
                s.append(n)
            if rx.search(re.split(r"\.\s", objeto, 1)[0][:200]):
                fuerte = True
    return s, fuerte

# --------------------------------------------------------------------------- BORME

ACTOS = [
    "Constitución", "Nombramientos", "Ceses/Dimisiones", "Revocaciones", "Reelecciones",
    "Ampliación de capital", "Reducción de capital", "Modificaciones estatutarias",
    "Cambio de domicilio social", "Cambio de objeto social", "Ampliacion del objeto social",
    "Ampliación del objeto social", "Cambio de denominación social", "Declaración de unipersonalidad",
    "Pérdida del caracter de unipersonalidad", "Pérdida del carácter de unipersonalidad",
    "Sociedad unipersonal", "Fusión por absorción", "Fusión por unión", "Escisión parcial", "Escisión total",
    "Segregación", "Cesión global de activo y pasivo", "Transformación de sociedad", "Disolución", "Extinción",
    "Situación concursal", "Otros conceptos", "Datos registrales", "Página web de la sociedad",
    "Desembolso de dividendos pasivos", "Emisión de obligaciones", "Depósito de libros",
    "Cierre provisional", "Reapertura hoja registral", "Fe de erratas", "Primera inscripción",
    "Adaptación Ley 2/95", "Adaptación Ley 44/2015", "Modificación de poderes", "Cancelaciones de oficio",
    "Crédito incobrable", "Reactivación de la sociedad", "Empresario Individual", "Adaptada segun D.T. 2 apartado 2 Ley 2/95",
    "Sucursal", "Acuerdo de ampliación de capital social sin ejecutar", "Unipersonal", "Socio único",
]
ACTO_RE = re.compile(r"(?<![\w/])(" + "|".join(re.escape(a) for a in sorted(ACTOS, key=len, reverse=True)) + r")\.", re.U)
CABECERA_RE = re.compile(r"^\s*(\d{5,7}) - (.+?)\.?\s*$")
RUIDO_RE = re.compile(r"^(BOLETÍN OFICIAL DEL REGISTRO MERCANTIL|Núm\. \d+|(Lunes|Martes|Miércoles|Jueves|Viernes|Sábado|Domingo) \d+ de .+ de \d{4}|Pág\. \d+|cve: BORME-.*|Verificable en .*|SECCIÓN PRIMERA|Empresarios|Actos inscritos|D\. L\.: .*|ISSN: .*|https?://\S+)$")


def borme_actos(texto):
    lineas = [l.rstrip() for l in texto.splitlines()]
    actos, cur = [], None
    for l in lineas:
        ls = l.strip()
        if not ls or RUIDO_RE.match(ls):
            continue
        m = CABECERA_RE.match(ls)
        if m and len(m.group(2)) > 1 and m.group(2).upper() == m.group(2) and not ls.endswith(":"):
            if cur:
                actos.append(cur)
            emp = re.sub(r"\s*\(R\.M\.[^)]*\)\s*$", "", m.group(2).strip())
            cur = {"num": m.group(1), "empresa": emp, "txt": ""}
            continue
        if cur is not None:
            sep = "" if cur["txt"].endswith("-") else " "
            cur["txt"] += sep + ls
    if cur:
        actos.append(cur)
    return actos


def trozos(txt):
    """'Constitución. Comienzo... Nombramientos. Adm. Unico: X.' -> {acto: contenido}"""
    pos = [(m.start(), m.end(), m.group(1)) for m in ACTO_RE.finditer(txt)]
    out = {}
    for i, (a, b, nombre) in enumerate(pos):
        fin = pos[i + 1][0] if i + 1 < len(pos) else len(txt)
        out.setdefault(nombre, []).append(txt[b:fin].strip(" ."))
    return out


def campo(txt, etiqueta, siguientes):
    m = re.search(re.escape(etiqueta) + r"\s*(.+?)(?=\s*(?:" + "|".join(re.escape(s) for s in siguientes) + r")|$)", txt)
    return m.group(1).strip(" .") if m else ""


def personas(nomb):
    """'Adm. Unico: GARCIA LUIS. Apoderado: X;Y.' -> [('Adm. Unico','GARCIA LUIS'),...]"""
    res = []
    for m in re.finditer(r"([A-ZÁÉÍÓÚ][\w\.\s/]{1,40}?):\s*([^:]+?)(?=\.\s+[A-ZÁÉÍÓÚ][\w\.\s/]{1,40}?:|\.?$)", nomb):
        cargo = m.group(1).strip()
        for p in m.group(2).split(";"):
            p = p.strip(" .")
            if p:
                res.append((cargo, p))
    return res


ES_EMPRESA = re.compile(r"\b(S\.?L\.?U?|S\.?A\.?U?|SOCIEDAD|SCCL|S\.? ?COOP|GMBH|B\.?V\.?|LTD|LLC|SARL|S\.?R\.?L|SPA|AG|FUND|FONDO|SOCIMI|GROUP|GRUPO|HOLDING)\b", re.I)


def clasificar_acto(a, provincia, fecha, pdf_url):
    t = trozos(a["txt"])
    nombre = a["empresa"]
    filas = []
    base = {"fecha": fecha, "provincia": provincia.title(), "empresa": nombre, "num_acto": a["num"], "fuente_url": pdf_url}

    objeto = domicilio = ""
    capital = None
    if "Constitución" in t:
        c = t["Constitución"][0]
        objeto = campo(c, "Objeto social:", ["Domicilio:", "Capital:"])
        domicilio = campo(c, "Domicilio:", ["Capital:"])
        capital = euros(campo(c, "Capital:", ["Declaración", "Nombramientos", "Datos registrales"]) or c)
    for k in ("Cambio de objeto social", "Ampliacion del objeto social", "Ampliación del objeto social"):
        if k in t and not objeto:
            objeto = t[k][0]
    if "Cambio de domicilio social" in t and not domicilio:
        domicilio = t["Cambio de domicilio social"][0]
    web = ""
    if "Página web de la sociedad" in t:
        web = t["Página web de la sociedad"][0].split()[0] if t["Página web de la sociedad"][0] else ""

    sect, fuerte = sectores_de(objeto, nombre)
    admins = []
    for blk in t.get("Nombramientos", []):
        admins += personas(blk)
    socio = ""
    for k in ("Socio único", "Declaración de unipersonalidad"):
        for blk in t.get(k, []):
            m = re.search(r"Socio [úu]nico:\s*(.+)", blk)
            socio = (m.group(1) if m else blk).strip(" .")
    if "Sociedad unipersonal" in t and not socio:
        pass

    amp = None
    for blk in t.get("Ampliación de capital", []):
        cap = euros(campo(blk, "Capital:", ["Resultante"]) or blk)
        res = euros(campo(blk, "Resultante Suscrito:", ["Datos", "."]))
        amp = (cap, res)

    señales = []
    if "Constitución" in t:
        señales.append("Constitución")
    if amp:
        señales.append("Ampliación de capital")
    if objeto and "Constitución" not in t:
        señales.append("Cambio/ampliación de objeto")
    if "Cambio de domicilio social" in t:
        señales.append("Cambio de domicilio")
    if admins and "Constitución" not in t:
        señales.append("Nombramientos")
    if any(k in t for k in ("Fusión por absorción", "Escisión parcial", "Escisión total", "Segregación")):
        señales.append("Fusión/escisión")
    if web:
        señales.append("Publica web")
    if not señales:
        return []
    disolucion = any(k in t for k in ("Disolución", "Extinción", "Situación concursal"))
    if disolucion:
        return []

    socio_empresa = bool(socio and ES_EMPRESA.search(socio))
    amp_cap = amp[0] if amp and amp[0] else 0
    # --- sin sector: solo ampliaciones grandes (inversión fuerte en cualquier sector)
    if not sect:
        if amp_cap >= 3_000_000:
            sect = ["Otro (ampliación grande)"]
        else:
            return []

    # --- prioridad (A solo para los sectores núcleo de RLE y con sector claro)
    NUCLEO = {"Promoción inmobiliaria", "Industrial y logística", "Energía y renovables", "Centros de datos",
              "Inversión inmobiliaria", "Hoteles y turismo", "Residencias y sanitario"}
    nucleo = fuerte and any(x in NUCLEO for x in sect)
    p = "C"
    if not fuerte and amp_cap < 500_000 and not web:
        p = "C"
    elif nucleo and ((amp_cap >= 500_000) or ("Constitución" in t and (capital or 0) >= 100_000)
                     or (socio_empresa and "Constitución" in t) or web or "Centros de datos" in sect):
        p = "A"
    elif amp_cap >= 3_000_000 or ("Constitución" in t and (capital or 0) >= 10_000) or amp_cap >= 60_000 or (socio_empresa and nucleo) or "Cambio/ampliación de objeto" in señales or "Fusión/escisión" in señales or "Nombramientos" in señales:
        p = "B"

    fila = dict(base)
    fila.update({
        "señales": ", ".join(señales),
        "sectores": ", ".join(sect),
        "sector_claro": "Sí" if fuerte else "",
        "prioridad": p,
        "capital_constitucion": capital,
        "ampliacion": amp_cap or None,
        "capital_resultante": amp[1] if amp else None,
        "objeto_social": objeto[:600],
        "domicilio": domicilio[:250],
        "socio_unico": socio[:200],
        "socio_es_empresa": "Sí" if socio_empresa else "",
        "administradores": "; ".join(f"{c}: {n}" for c, n in admins)[:400],
        "web_borme": web,
        "texto": a["txt"][:900],
        "clave": norm_nombre(nombre),
    })
    filas.append(fila)
    return filas


def borme_sumario(fecha):
    ymd = fecha.replace("-", "")
    raw = http_get(f"https://www.boe.es/datosabiertos/api/borme/sumario/{ymd}", accept="application/json")
    data = json.loads(raw)
    items = []
    for d in walk(data):
        ident = d.get("identificador") or d.get("id") or ""
        if isinstance(ident, str) and ident.startswith("BORME-A-"):
            url = as_text(d.get("url_pdf"))
            titulo = as_text(d.get("titulo"))
            if url:
                items.append({"id": ident, "provincia": titulo, "url": url})
    seen, out = set(), []
    for i in items:
        if i["url"] not in seen:
            seen.add(i["url"]); out.append(i)
    return out


def cmd_borme(fecha, salida):
    try:
        items = borme_sumario(fecha)
    except urllib.error.HTTPError as e:
        if e.code == 404:
            json.dump({"fecha": fecha, "estado": "sin BORME ese día", "filas": []}, open(salida, "w"), ensure_ascii=False)
            print("Sin BORME para", fecha); return
        raise
    filas, errores, total_actos = [], [], 0
    for it in items:
        try:
            txt = pdf_texto(http_get(it["url"]))
            actos = borme_actos(txt)
            total_actos += len(actos)
            for a in actos:
                filas += clasificar_acto(a, it["provincia"], fecha, it["url"])
        except Exception as e:
            errores.append(f"{it['id']}: {e}")
        time.sleep(0.3)
    res = {"fecha": fecha, "estado": "ok", "pdfs": len(items), "actos_leidos": total_actos,
           "errores": errores, "filas": filas}
    json.dump(res, open(salida, "w"), ensure_ascii=False, indent=1)
    from collections import Counter
    print(f"BORME {fecha}: {len(items)} PDFs, {total_actos} actos, {len(filas)} señales "
          f"{dict(Counter(f['prioridad'] for f in filas))}, errores {len(errores)}")

# --------------------------------------------------------------------------- BOE V-B renovables

RENOV_RE = re.compile(r"fotovolt|solar|e[oó]lic|almacenamiento|bater[ií]as|hibridaci|hidr[oó]geno|biog[aá]s|biometano|electrolizador|parque e[oó]lico|planta de generaci", re.I)
PROMOTOR_RE = re.compile(
    r"(?:promovid[oa]s?(?: por)?|peticionari[oa]s?|solicitante|promotor(?:a)?|titular)\s*[:,]?\s*(?:la (?:mercantil|sociedad|empresa)\s*)?[«\"“]?"
    r"([A-ZÁÉÍÓÚÑ0-9][A-Za-zÁÉÍÓÚÑáéíóúñ0-9&\.,'’\- ]{2,90}?(?:S\.?\s?L\.?\s?U?\.?|S\.?\s?A\.?\s?U?\.?|SOCIEDAD LIMITADA(?: UNIPERSONAL)?|SOCIEDAD AN[ÓO]NIMA(?: UNIPERSONAL)?|S\.?L\.?L\.?))(?=[\s,\.;\)»\"”]|$)")
MW_RE = re.compile(r"(\d{1,4}(?:[\.,]\d{1,3})?)\s*(MW|MWp|MWn|MVA|kWp)", re.I)


def cmd_boe5b(fecha, salida):
    ymd = fecha.replace("-", "")
    try:
        data = json.loads(http_get(f"https://www.boe.es/datosabiertos/api/boe/sumario/{ymd}", accept="application/json"))
    except urllib.error.HTTPError as e:
        if e.code == 404:
            json.dump({"fecha": fecha, "estado": "sin BOE", "filas": []}, open(salida, "w")); return
        raise
    cand = []
    for d in walk(data):
        if isinstance(d, dict) and str(d.get("codigo", "")).upper() in ("5B", "5V"):
            for it in walk(d):
                ident = it.get("identificador")
                tit = as_text(it.get("titulo"))
                if ident and tit and RENOV_RE.search(tit):
                    cand.append({"id": ident, "titulo": tit, "html": as_text(it.get("url_html")),
                                 "xml": as_text(it.get("url_xml")), "pdf": as_text(it.get("url_pdf"))})
    # algunos sumarios no marcan la sección en el nodo padre: respaldo por identificador BOE-B
    if not cand:
        for it in walk(data):
            ident = it.get("identificador") or ""
            tit = as_text(it.get("titulo"))
            if isinstance(ident, str) and ident.startswith("BOE-B-") and tit and RENOV_RE.search(tit):
                cand.append({"id": ident, "titulo": tit, "html": as_text(it.get("url_html")),
                             "xml": as_text(it.get("url_xml")), "pdf": as_text(it.get("url_pdf"))})
    filas = []
    vistos = set()
    for c in cand:
        if c["id"] in vistos:
            continue
        vistos.add(c["id"])
        texto = c["titulo"]
        try:
            if c["xml"]:
                root = ET.fromstring(http_get(c["xml"]))
                texto = c["titulo"] + " " + " ".join(t.strip() for t in root.itertext() if t.strip())
            time.sleep(0.3)
        except Exception:
            pass
        proms = []
        for m in PROMOTOR_RE.finditer(texto):
            p = re.sub(r"\s+", " ", m.group(1)).strip(" ,.")
            if p.upper() not in [x.upper() for x in proms] and len(p) < 100:
                proms.append(p)
        mw = MW_RE.search(texto)
        tecno = [k for k, rx in [("Fotovoltaica", r"fotovolt|solar"), ("Eólica", r"e[oó]lic"),
                                 ("Almacenamiento/BESS", r"almacenamiento|bater"), ("Hidrógeno", r"hidr[oó]geno|electroliz"),
                                 ("Biogás/biometano", r"biog|biomet")] if re.search(rx, texto, re.I)]
        prov = ""
        m = re.search(r"(?:Subdelegaci[oó]n|Delegaci[oó]n) del Gobierno en ([A-ZÁÉÍÓÚ][\wáéíóúñ\- ]+?)(?: por| sobre| relativ|,|\.)", texto)
        if m:
            prov = m.group(1).strip()
        filas.append({"fecha": fecha, "id_boe": c["id"], "titulo": c["titulo"][:500],
                      "promotor": proms[0] if proms else "", "otros_promotores": "; ".join(proms[1:4]),
                      "tecnologia": ", ".join(tecno), "potencia": (mw.group(1) + " " + mw.group(2)) if mw else "",
                      "provincia": prov, "url": c["html"] or c["pdf"],
                      "clave": norm_nombre(proms[0]) if proms else ""})
    json.dump({"fecha": fecha, "estado": "ok", "anuncios": len(filas), "filas": filas}, open(salida, "w"), ensure_ascii=False, indent=1)
    print(f"BOE V-B {fecha}: {len(filas)} anuncios de renovables, {sum(1 for f in filas if f['promotor'])} con promotor")

# --------------------------------------------------------------------------- PLACSP

PLACSP_FEED = "https://contrataciondelsectorpublico.gob.es/sindicacion/sindicacion_643/licitacionesPerfilesContratanteCompleto3.atom"
NO_SUELO_RE = re.compile(r"madera|forestal|aprovechamiento|chopo|pinar|le[ñn]a|veh[ií]culo|cisterna|chatarra|maquinaria|tractor|ganado|caza|pastos|corcho|resina|mobiliario|equipos", re.I)
SUELO_RE = re.compile(r"enajenaci[oó]n.{0,80}(parcela|solar|finca|suelo|terreno|inmueble|edificio|local|nave|vivienda|bien patrimonial)|venta de (la |las |los )?(parcela|solar|finca|suelo|terreno|inmueble)|subasta.{0,40}(parcela|solar|finca|suelo|terreno)|derecho de superficie|permuta|concurso.{0,40}(parcela|suelo)", re.I)
EDIF_RE = re.compile(r"nave|log[ií]stic|industrial|vivienda|edificio|residencia|hotel|aparcamiento|centro de datos|urbanizaci[oó]n", re.I)


def local(tag):
    return tag.rsplit("}", 1)[-1]


def find_all(el, name):
    return [e for e in el.iter() if local(e.tag) == name]


def first_text(el, *path):
    cur = [el]
    for name in path:
        nxt = []
        for c in cur:
            nxt += [e for e in c.iter() if local(e.tag) == name]
        cur = nxt
        if not cur:
            return ""
    return (cur[0].text or "").strip()


def cmd_placsp(fecha, salida, max_paginas=60):
    fin = datetime.fromisoformat(fecha).replace(hour=23, minute=59, tzinfo=timezone.utc)
    desde = fin - timedelta(days=1, hours=12)
    url, pag, filas, vistos, leidas = PLACSP_FEED, 0, [], set(), 0
    parar = False
    while url and pag < max_paginas and not parar:
        root = ET.fromstring(http_get(url, timeout=120))
        pag += 1
        nxt = ""
        for l in root:
            if local(l.tag) == "link" and l.get("rel") == "next":
                nxt = l.get("href")
        for e in root:
            if local(e.tag) != "entry":
                continue
            leidas += 1
            upd = first_text(e, "updated")
            try:
                t = datetime.fromisoformat(upd.replace("Z", "+00:00"))
            except Exception:
                t = None
            if t and t < desde:
                parar = True
                continue
            if t and t > fin + timedelta(days=1):
                continue
            estado = first_text(e, "ContractFolderStatusCode")
            if estado not in ("ADJ", "RES"):
                continue
            nombre = first_text(e, "ProcurementProject", "Name") or first_text(e, "title")
            tipo = first_text(e, "ProcurementProject", "TypeCode")
            importe = first_text(e, "ProcurementProject", "BudgetAmount", "TaxExclusiveAmount") or first_text(e, "TotalAmount")
            try:
                imp = float(importe)
            except Exception:
                imp = 0.0
            es_suelo = bool(SUELO_RE.search(nombre)) and not NO_SUELO_RE.search(nombre)
            es_obra = tipo == "3" and imp >= 1_500_000 and bool(EDIF_RE.search(nombre))
            if not (es_suelo or es_obra):
                continue
            organo = first_text(e, "LocatedContractingParty", "PartyName", "Name")
            lugar = first_text(e, "RealizedLocation", "CountrySubentity")
            link = ""
            for l in e:
                if local(l.tag) == "link":
                    link = l.get("href") or link
            for tr in find_all(e, "TenderResult"):
                for wp in find_all(tr, "WinningParty"):
                    gan = first_text(wp, "PartyName", "Name")
                    nif = first_text(wp, "PartyIdentification", "ID")
                    if not gan:
                        continue
                    k = (norm_nombre(gan), nombre[:80])
                    if k in vistos:
                        continue
                    vistos.add(k)
                    adj_imp = first_text(tr, "AwardedTenderedProject", "LegalMonetaryTotal", "TaxExclusiveAmount")
                    filas.append({"fecha": fecha, "tipo": "Suelo público adjudicado" if es_suelo else "Obra de edificación adjudicada",
                                  "prioridad": "A" if es_suelo else "B", "empresa": gan, "nif": nif, "objeto": nombre[:400],
                                  "organo": organo, "lugar": lugar, "importe_licitacion": imp or None,
                                  "importe_adjudicacion": float(adj_imp) if adj_imp.replace('.', '', 1).isdigit() else None,
                                  "url": link, "clave": norm_nombre(gan)})
        url = nxt
        time.sleep(0.5)
    json.dump({"fecha": fecha, "estado": "ok", "paginas": pag, "entradas_leidas": leidas, "filas": filas},
              open(salida, "w"), ensure_ascii=False, indent=1)
    print(f"PLACSP: {pag} páginas, {leidas} entradas, {len(filas)} adjudicatarias de interés")

# --------------------------------------------------------------------------- RSS de prensa sectorial

FEEDS = [
    ("Brainsre (inmobiliario)", "https://brainsre.news/feed/"),
    ("EjePrime (inmobiliario)", "https://www.ejeprime.com/rss"),
    ("Interempresas Logística", "https://www.interempresas.net/Logistica/RSS/"),
    ("Cadena de Suministro", "https://www.cadena-de-suministro.es/feed/"),
    ("pv magazine España", "https://www.pv-magazine.es/feed/"),
    ("Energías Renovables", "https://www.energias-renovables.com/rss"),
    ("elEconomista Inmobiliario", "https://www.eleconomista.es/rss/rss-ecoley-inmobiliario.php"),
    ("Expansión Inmobiliario", "https://e00-expansion.uecdn.es/rss/empresasinmobiliario.xml"),
    ("Cinco Días Compañías", "https://cincodias.elpais.com/seccion/rss/companias/"),
]
SEÑAL_RE = re.compile(r"compra|adquier|vend|suelo|solar|parcela|terreno|nave|log[ií]stic|inversi|invert|abrir[aá]|abre|nuev[oa]s? (centro|planta|f[aá]brica|sede|hotel|residencia)|expansi|ampl[ií]a|construir[aá]|promoci|viviendas|socimi|fondo|family office|data ?cent|centro de datos|fotovolt|e[oó]lic|bater|hidr[oó]geno|biometano|MW\b|m2|m²|millones", re.I)


def cmd_rss(fecha, salida, horas=40):
    from email.utils import parsedate_to_datetime
    lim = datetime.fromisoformat(fecha).replace(hour=23, tzinfo=timezone.utc) - timedelta(hours=horas)
    items, estado = [], {}
    for nombre, url in FEEDS:
        try:
            root = ET.fromstring(http_get(url, timeout=40, retries=2))
            n = 0
            for it in root.iter():
                if local(it.tag) not in ("item", "entry"):
                    continue
                tit = first_text(it, "title")
                link = first_text(it, "link")
                if not link:
                    for l in it:
                        if local(l.tag) == "link" and l.get("href"):
                            link = l.get("href"); break
                fecha_txt = first_text(it, "pubDate") or first_text(it, "updated") or first_text(it, "published")
                try:
                    t = parsedate_to_datetime(fecha_txt) if "," in fecha_txt else datetime.fromisoformat(fecha_txt.replace("Z", "+00:00"))
                    if t.tzinfo is None:
                        t = t.replace(tzinfo=timezone.utc)
                except Exception:
                    t = None
                if t and t < lim:
                    continue
                desc = re.sub(r"<[^>]+>", " ", first_text(it, "description") or first_text(it, "summary"))
                desc = re.sub(r"\s+", " ", desc).strip()[:400]
                if not SEÑAL_RE.search(tit + " " + desc):
                    continue
                items.append({"medio": nombre, "titulo": tit, "resumen": desc, "url": link,
                              "fecha": t.date().isoformat() if t else ""})
                n += 1
            estado[nombre] = f"ok ({n})"
        except Exception as e:
            estado[nombre] = f"ERROR {type(e).__name__}"
        time.sleep(0.3)
    json.dump({"fecha": fecha, "feeds": estado, "items": items}, open(salida, "w"), ensure_ascii=False, indent=1)
    print(f"RSS: {len(items)} titulares con señal · " + "; ".join(f"{k}: {v}" for k, v in estado.items()))

# --------------------------------------------------------------------------- main

if __name__ == "__main__":
    cmd = sys.argv[1]
    if cmd == "borme":
        cmd_borme(sys.argv[2], sys.argv[3])
    elif cmd == "boe5b":
        cmd_boe5b(sys.argv[2], sys.argv[3])
    elif cmd == "placsp":
        cmd_placsp(sys.argv[2], sys.argv[3])
    elif cmd == "rss":
        cmd_rss(sys.argv[2], sys.argv[3])
    elif cmd == "check":
        for u in ["https://www.boe.es/datosabiertos/api/borme/sumario/" + sys.argv[2].replace("-", ""), PLACSP_FEED] + [f[1] for f in FEEDS]:
            try:
                http_get(u, accept="application/json" if "datosabiertos" in u else None, retries=1, timeout=20); print("OK   ", u)
            except Exception as e:
                print("FALLA", u, type(e).__name__, str(e)[:80])
    elif cmd == "parse-pdf":
        txt = pdf_texto(open(sys.argv[2], "rb").read())
        filas = []
        actos = borme_actos(txt)
        for a in actos:
            filas += clasificar_acto(a, sys.argv[3] if len(sys.argv) > 3 else "?", "test", sys.argv[2])
        print(len(actos), "actos;", len(filas), "señales")
        for f in filas:
            print(json.dumps({k: f[k] for k in ("empresa", "señales", "sectores", "prioridad", "ampliacion", "capital_constitucion", "objeto_social", "domicilio", "socio_unico", "administradores")}, ensure_ascii=False))
    else:
        print(__doc__)
