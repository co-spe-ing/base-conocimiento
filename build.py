#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generador de la Base de Conocimiento Jurídica DIAN.

Convierte los PDF y DOCX que estén en esta misma carpeta en páginas HTML con el
texto íntegro visible, pensadas para que Bing las indexe y el agente de Copilot
pueda leerlas (Copilot NO abre PDF enlazados: solo lee el texto del HTML).

Uso:
    python3 build.py

Requisitos:
    - poppler-utils (pdftotext, pdfinfo)   ->  apt install poppler-utils
    - pandoc                               ->  apt install pandoc

Salida: archivos .html planos en esta carpeta, más sitemap.xml, robots.txt,
indice-tematico.html e index.html.
"""

import os
import re
import html
import json
import glob
import shutil
import subprocess
import unicodedata
from collections import Counter
from datetime import date

# --------------------------------------------------------------------------
# CONFIGURACIÓN
# --------------------------------------------------------------------------

SITIO = "https://co-spe-ing.github.io/base-conocimiento"
TITULO_SITIO = "Base de Conocimiento Jurídica DIAN"
ENTIDAD = "Dirección de Impuestos y Aduanas Nacionales (DIAN)"
HOY = date.today().isoformat()

# Tamaño objetivo de cada página en caracteres. Bing devuelve fragmentos cortos:
# páginas demasiado largas producen respuestas ruidosas; demasiado cortas
# multiplican las URLs que hay que indexar.
OBJETIVO_CHARS = 38000
MAX_CHARS = 60000

# Ficha de cada documento: archivo -> metadatos.
# "alias" son las formas en que la gente realmente busca la norma: se imprimen
# en la página para que Bing las asocie al documento.
DOCUMENTOS = [
    {
        "archivo": "Ley_909_de_2004.pdf",
        "slug": "ley-909-de-2004",
        "titulo": "Ley 909 de 2004",
        "subtitulo": "Por la cual se expiden normas que regulan el empleo público, la carrera administrativa, gerencia pública y se dictan otras disposiciones",
        "categoria": "Normativa",
        "tipo": "Ley",
        "emisor": "Congreso de la República de Colombia",
        "fecha": "2004-09-23",
        "alias": ["Ley 909", "ley de carrera administrativa", "ley de empleo público",
                  "Ley 909 04", "estatuto de carrera administrativa"],
    },
    {
        "archivo": "Decreto_1083_de_2015_Sector_de_Función_Pública.pdf",
        "slug": "decreto-1083-de-2015",
        "titulo": "Decreto 1083 de 2015",
        "subtitulo": "Decreto Único Reglamentario del Sector de Función Pública",
        "categoria": "Normativa",
        "tipo": "Decreto único reglamentario",
        "emisor": "Departamento Administrativo de la Función Pública",
        "fecha": "2015-05-26",
        "alias": ["Decreto 1083", "DURSFP", "decreto único de función pública",
                  "decreto reglamentario de la Ley 909"],
    },
    {
        "archivo": "Decreto_927_de_2023.pdf",
        "slug": "decreto-ley-927-de-2023",
        "titulo": "Decreto Ley 927 de 2023",
        "subtitulo": "Sistema Específico de Carrera Administrativa de la DIAN",
        "categoria": "Normativa",
        "tipo": "Decreto Ley",
        "emisor": "Presidencia de la República de Colombia",
        "fecha": "2023-06-07",
        "alias": ["Decreto 927", "Decreto 927 de 2023", "carrera DIAN",
                  "sistema específico de carrera DIAN", "decreto ley 927"],
    },
    {
        "archivo": "Ley_1437_de_2011.pdf",
        "slug": "ley-1437-de-2011-cpaca",
        "titulo": "Ley 1437 de 2011 (CPACA)",
        "subtitulo": "Código de Procedimiento Administrativo y de lo Contencioso Administrativo",
        "categoria": "Normativa",
        "tipo": "Ley / Código",
        "emisor": "Congreso de la República de Colombia",
        "fecha": "2011-01-18",
        "alias": ["CPACA", "Ley 1437", "código de procedimiento administrativo",
                  "derecho de petición", "silencio administrativo", "recurso de reposición"],
    },
    {
        "archivo": "Ley_1712_de_2014.pdf",
        "slug": "ley-1712-de-2014",
        "titulo": "Ley 1712 de 2014",
        "subtitulo": "Ley de Transparencia y del Derecho de Acceso a la Información Pública Nacional",
        "categoria": "Normativa",
        "tipo": "Ley",
        "emisor": "Congreso de la República de Colombia",
        "fecha": "2014-03-06",
        "alias": ["Ley 1712", "ley de transparencia", "acceso a la información pública",
                  "información reservada", "información clasificada"],
    },
    {
        "archivo": "Constitución_Política_1_de_1991_Asamblea_Nacional_Constituyente.pdf",
        "slug": "constitucion-politica-1991",
        "titulo": "Constitución Política de Colombia de 1991",
        "subtitulo": "Texto vigente con concordancias",
        "categoria": "Normativa",
        "tipo": "Constitución",
        "emisor": "Asamblea Nacional Constituyente",
        "fecha": "1991-07-04",
        "alias": ["Constitución Política", "Constitución de 1991", "carta política",
                  "artículo 23 derecho de petición", "artículo 125 carrera administrativa"],
    },
    {
        "archivo": "Resolución 000067 de 11042024.pdf",
        "slug": "resolucion-000067-de-2024",
        "titulo": "Resolución 000067 del 11 de abril de 2024",
        "subtitulo": "Resolución DIAN",
        "categoria": "Normativa",
        "tipo": "Resolución",
        "emisor": "Dirección de Impuestos y Aduanas Nacionales (DIAN)",
        "fecha": "2024-04-11",
        "alias": ["Resolución 67 de 2024", "Resolución 000067", "resolución DIAN 2024"],
    },
    {
        "archivo": "Proteccion_Datos_DIAN.pdf",
        "slug": "proteccion-de-datos-dian",
        "titulo": "Protección de datos personales — DIAN",
        "subtitulo": "Lineamientos institucionales de tratamiento de datos personales",
        "categoria": "Normativa",
        "tipo": "Documento institucional",
        "emisor": "Dirección de Impuestos y Aduanas Nacionales (DIAN)",
        "fecha": "",
        "alias": ["habeas data", "tratamiento de datos personales DIAN",
                  "política de protección de datos", "Ley 1581"],
    },
    {
        "archivo": "C-197 2025.pdf",
        "slug": "sentencia-c-197-de-2025",
        "titulo": "Sentencia C-197 de 2025",
        "subtitulo": "Corte Constitucional — control de constitucionalidad",
        "categoria": "Jurisprudencia",
        "tipo": "Sentencia de constitucionalidad",
        "emisor": "Corte Constitucional de Colombia",
        "fecha": "2025",
        "alias": ["C-197/25", "C-197 de 2025", "sentencia C 197"],
    },
    {
        "archivo": "SL2600-2025.pdf",
        "slug": "sentencia-sl2600-de-2025",
        "titulo": "Sentencia SL2600-2025",
        "subtitulo": "Corte Suprema de Justicia — Sala de Casación Laboral",
        "categoria": "Jurisprudencia",
        "tipo": "Sentencia de casación laboral",
        "emisor": "Corte Suprema de Justicia, Sala de Casación Laboral",
        "fecha": "2025",
        "alias": ["SL2600", "SL2600-2025", "sentencia laboral 2600 de 2025"],
    },
    {
        "archivo": "T-084 2018.pdf",
        "slug": "sentencia-t-084-de-2018",
        "titulo": "Sentencia T-084 de 2018",
        "subtitulo": "Corte Constitucional — acción de tutela",
        "categoria": "Jurisprudencia",
        "tipo": "Sentencia de tutela",
        "emisor": "Corte Constitucional de Colombia",
        "fecha": "2018",
        "alias": ["T-084/18", "T-084 de 2018", "sentencia T 084"],
    },
    {
        "archivo": "sentencia t-737-2017.pdf",
        "slug": "sentencia-t-737-de-2017",
        "titulo": "Sentencia T-737 de 2017",
        "subtitulo": "Corte Constitucional — acción de tutela",
        "categoria": "Jurisprudencia",
        "tipo": "Sentencia de tutela",
        "emisor": "Corte Constitucional de Colombia",
        "fecha": "2017",
        "alias": ["T-737/17", "T-737 de 2017", "sentencia T 737"],
    },
    {
        "archivo": "PLANTILLA MODIFICA NOMBRAMIENTO CAMBIO DE SEDE.docx",
        "slug": "plantilla-modificacion-nombramiento-cambio-de-sede",
        "titulo": "Plantilla — Modificación de nombramiento por cambio de sede",
        "subtitulo": "Modelo de acto administrativo",
        "categoria": "Plantillas",
        "tipo": "Plantilla / modelo de acto administrativo",
        "emisor": "DIAN",
        "fecha": "",
        "alias": ["modificar nombramiento", "cambio de sede", "reubicación de sede",
                  "modelo resolución cambio de sede"],
    },
    {
        "archivo": "PLANTILLA RESPUESTA A PETICION DE SECCIONALES.docx",
        "slug": "plantilla-respuesta-peticion-seccionales",
        "titulo": "Plantilla — Respuesta a petición de seccionales",
        "subtitulo": "Modelo de respuesta a derecho de petición",
        "categoria": "Plantillas",
        "tipo": "Plantilla / modelo de oficio",
        "emisor": "DIAN",
        "fecha": "",
        "alias": ["responder derecho de petición", "modelo respuesta petición",
                  "oficio de respuesta seccional"],
    },
    {
        "archivo": "PLANTILLA RESUELVE RECURSO REUBICACION CONFIRMANDO.docx",
        "slug": "plantilla-resuelve-recurso-reubicacion-confirmando",
        "titulo": "Plantilla — Resolución de recurso de reubicación (confirma)",
        "subtitulo": "Modelo de acto que resuelve recurso de reposición confirmando la decisión",
        "categoria": "Plantillas",
        "tipo": "Plantilla / modelo de acto administrativo",
        "emisor": "DIAN",
        "fecha": "",
        "alias": ["recurso de reposición reubicación", "confirmar decisión reubicación",
                  "modelo resolución recurso"],
    },
]

# Índice temático: tema -> palabras clave que se buscan en los epígrafes de los
# artículos y en los títulos de sección.
TEMAS = {
    "Derecho de petición y respuesta a peticiones": [
        "petición", "peticiones", "peticionario", "consulta", "queja", "reclamo",
        "término para resolver", "silencio administrativo"],
    "Carrera administrativa y concursos": [
        "carrera administrativa", "concurso", "mérito", "lista de elegibles",
        "convocatoria", "período de prueba", "ingreso", "selección"],
    "Nombramiento, posesión y retiro": [
        "nombramiento", "posesión", "retiro", "insubsistencia", "renuncia",
        "vacancia", "supresión", "provisional", "encargo"],
    "Traslado, reubicación y cambio de sede": [
        "traslado", "reubicación", "reubicar", "sede", "ubicación", "planta global",
        "comisión de servicios", "movimiento"],
    "Situaciones administrativas y permisos": [
        "situación administrativa", "licencia", "permiso", "vacaciones",
        "comisión", "encargo", "servicio activo"],
    "Evaluación del desempeño": [
        "evaluación del desempeño", "calificación", "desempeño", "acuerdo de gestión"],
    "Régimen disciplinario y responsabilidad": [
        "disciplinario", "falta", "sanción", "responsabilidad", "inhabilidad",
        "incompatibilidad", "conflicto de interés"],
    "Procedimiento administrativo, notificaciones y recursos": [
        "notificación", "comunicación", "recurso", "reposición", "apelación",
        "revocatoria", "firmeza", "acto administrativo", "procedimiento administrativo"],
    "Transparencia, acceso a la información y datos personales": [
        "información pública", "transparencia", "reserva", "clasificada",
        "datos personales", "habeas data", "publicidad", "archivo"],
    "Derechos fundamentales y garantías constitucionales": [
        "derecho fundamental", "tutela", "debido proceso", "igualdad",
        "trabajo", "dignidad", "estabilidad"],
    "Sistema específico de carrera de la DIAN": [
        "DIAN", "impuestos", "aduanas", "específico de carrera", "gestión tributaria"],
}

LIGADURAS = {
    "ﬀ": "ff", "ﬁ": "fi", "ﬂ": "fl", "ﬃ": "ffi", "ﬄ": "ffl",
    "ﬅ": "st", "ﬆ": "st", "­": "",
    "‘": "'", "’": "'", "“": '"', "”": '"',
    "–": "-", "—": "—", " ": " ",
}

RE_ARTICULO = re.compile(
    r"^(ART[IÍ]CULO|Art[ií]culo|ARTICULO)\s+"
    # Admite numeración simple (12), con letra (12A) y decimal (2.2.2.4.6),
    # que es la que usa el Decreto Único 1083 de 2015.
    r"(\d+(?:\.\d+)+|\d+[A-Za-z]?(?:\s*[°º])?(?:\s*[-–]\s*\d+)?)"
    r"\s*[.\-–:]?\s*(.*)$")
RE_SECCION = re.compile(
    r"^(LIBRO|PARTE|T[IÍ]TULO|TITULO|CAP[IÍ]TULO|CAPITULO|SECCI[OÓ]N|SECCION|SUBSECCI[OÓ]N)"
    r"\s+([A-Z0-9IVXLC]+[.\-]?)\s*(.*)$", re.IGNORECASE)


# --------------------------------------------------------------------------
# UTILIDADES
# --------------------------------------------------------------------------

def slugify(texto, maxlen=60):
    t = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode()
    t = re.sub(r"[^a-zA-Z0-9]+", "-", t).strip("-").lower()
    return t[:maxlen].strip("-") or "seccion"


def esc(s):
    return html.escape(s, quote=True)


def normalizar(texto):
    for k, v in LIGADURAS.items():
        texto = texto.replace(k, v)
    return texto


def extraer_texto(ruta):
    """Devuelve el texto plano del documento, página a página."""
    ext = os.path.splitext(ruta)[1].lower()
    if ext == ".pdf":
        out = subprocess.run(["pdftotext", "-enc", "UTF-8", ruta, "-"],
                             capture_output=True, text=True)
        return normalizar(out.stdout).split("\f")
    if ext in (".docx", ".doc"):
        out = subprocess.run(["pandoc", "-t", "plain", "--wrap=none", ruta],
                             capture_output=True, text=True)
        return [normalizar(out.stdout)]
    raise ValueError("formato no soportado: " + ruta)


def quitar_encabezados_repetidos(paginas):
    """Elimina encabezados, pies y números de página que se repiten en el PDF.

    Una línea que aparece en más del 25 % de las páginas es plantilla, no
    contenido: si se deja, contamina cada fragmento que Bing devuelva.
    """
    if len(paginas) < 4:
        limpio = paginas
    else:
        cont = Counter()
        for p in paginas:
            for l in set(x.strip() for x in p.splitlines() if x.strip()):
                cont[l] += 1
        umbral = max(3, int(len(paginas) * 0.25))
        basura = {l for l, c in cont.items() if c >= umbral and len(l) < 200}
        limpio = []
        for p in paginas:
            lineas = [l for l in p.splitlines() if l.strip() not in basura]
            limpio.append("\n".join(lineas))
    texto = "\n".join(limpio)
    # Números de página sueltos y marcas de paginación.
    texto = re.sub(r"^\s*\d{1,4}\s*$", "", texto, flags=re.M)
    texto = re.sub(r"^\s*P[áa]gina\s+\d+\s*(de\s+\d+)?\s*$", "", texto, flags=re.M | re.I)
    return texto


def reflujo(texto):
    """Une los renglones cortados por el PDF y devuelve párrafos completos."""
    texto = re.sub(r"[ \t]+", " ", texto)
    bloques = re.split(r"\n\s*\n", texto)
    parrafos = []
    for b in bloques:
        lineas = [l.strip() for l in b.splitlines() if l.strip()]
        if not lineas:
            continue
        buf = ""
        for l in lineas:
            # Una sección es un rótulo aislado: se cierra en su propia línea.
            if RE_SECCION.match(l) and len(l) < 300:
                if buf:
                    parrafos.append(buf)
                    buf = ""
                parrafos.append(l)
                continue
            # Un artículo abre párrafo nuevo, pero sigue absorbiendo los
            # renglones que el PDF cortó a mitad de frase.
            if RE_ARTICULO.match(l) or not buf:
                if buf:
                    parrafos.append(buf)
                buf = l
                continue
            if buf.endswith("-"):
                buf = buf[:-1] + l
            else:
                buf = buf + " " + l
        if buf:
            parrafos.append(buf)
    return [p.strip() for p in parrafos if p.strip()]


# --------------------------------------------------------------------------
# ESTRUCTURA DEL DOCUMENTO
# --------------------------------------------------------------------------

def estructurar(parrafos):
    """Convierte la lista de párrafos en una lista de bloques tipados."""
    bloques = []
    for p in parrafos:
        m = RE_ARTICULO.match(p)
        if m and len(p) < 8000:
            num = re.sub(r"\s+", "", m.group(2)).rstrip(".°º")
            epigrafe = m.group(3).strip()
            cuerpo = ""
            # El epígrafe suele terminar en punto; lo que sigue ya es texto.
            if len(epigrafe) > 90:
                mm = re.match(r"^(.{0,120}?[.:])\s+(.*)$", epigrafe, re.S)
                if mm:
                    epigrafe, cuerpo = mm.group(1).strip(), mm.group(2).strip()
                else:
                    cuerpo, epigrafe = epigrafe, ""
            bloques.append({"tipo": "articulo", "num": num,
                            "epigrafe": epigrafe, "texto": cuerpo})
            continue
        m = RE_SECCION.match(p)
        if m and len(p) < 300:
            bloques.append({"tipo": "seccion", "texto": p.strip()})
            continue
        bloques.append({"tipo": "parrafo", "texto": p})
    return bloques


def paginar(bloques, objetivo=OBJETIVO_CHARS, maximo=MAX_CHARS):
    """Agrupa bloques en páginas, cortando en fronteras de sección o artículo."""
    paginas, actual, tam = [], [], 0
    for b in bloques:
        largo = len(b.get("texto", "")) + len(b.get("epigrafe", "")) + 40
        corte = b["tipo"] in ("seccion", "articulo")
        if actual and ((tam >= objetivo and corte) or tam + largo > maximo):
            paginas.append(actual)
            actual, tam = [], 0
        actual.append(b)
        tam += largo
    if actual:
        paginas.append(actual)
    return paginas


def etiqueta_pagina(bloques):
    """Título humano de la página: rango de artículos y/o sección."""
    arts = [b["num"] for b in bloques if b["tipo"] == "articulo"]
    secs = [b["texto"] for b in bloques if b["tipo"] == "seccion"]
    partes = []
    if secs:
        partes.append(secs[0][:90])
    if arts:
        if len(arts) == 1:
            partes.append("Artículo %s" % arts[0])
        else:
            partes.append("Artículos %s a %s" % (arts[0], arts[-1]))
    return " — ".join(partes) if partes else "Texto"


# --------------------------------------------------------------------------
# PLANTILLAS HTML
# --------------------------------------------------------------------------

CSS = """:root{--tinta:#16202c;--suave:#5b6b7c;--linea:#dfe5ec;--fondo:#fff;--acento:#0b4f8a;--caja:#f5f8fb}
*{box-sizing:border-box}
body{margin:0;background:var(--fondo);color:var(--tinta);
font:16px/1.65 Georgia,"Times New Roman",serif}
.envoltura{max-width:820px;margin:0 auto;padding:24px 20px 72px}
header.sitio{border-bottom:2px solid var(--linea);margin-bottom:28px;padding-bottom:14px}
header.sitio a{color:var(--acento);text-decoration:none;font:600 14px/1.4 system-ui,sans-serif;
letter-spacing:.04em;text-transform:uppercase}
nav.miga{font:13px/1.5 system-ui,sans-serif;color:var(--suave);margin:0 0 18px}
nav.miga a{color:var(--acento)}
h1{font-size:1.7rem;line-height:1.25;margin:.2em 0 .1em}
h2{font-size:1.18rem;margin:2.1em 0 .4em;padding-top:.2em;border-top:1px solid var(--linea)}
h3{font-size:1.02rem;margin:1.6em 0 .3em;color:var(--suave);
font:600 .95rem/1.4 system-ui,sans-serif;text-transform:uppercase;letter-spacing:.05em}
p{margin:.7em 0;text-align:justify}
.sub{color:var(--suave);font-size:1rem;margin:.2em 0 1em}
.ficha{background:var(--caja);border:1px solid var(--linea);border-radius:6px;
padding:14px 18px;margin:18px 0;font:14px/1.6 system-ui,sans-serif}
.ficha dl{display:grid;grid-template-columns:auto 1fr;gap:2px 14px;margin:0}
.ficha dt{color:var(--suave)}
.ficha dd{margin:0}
.aviso{font:13px/1.6 system-ui,sans-serif;color:var(--suave);
border-left:3px solid var(--linea);padding:2px 0 2px 14px;margin:22px 0}
ul.lista{font:15px/1.7 system-ui,sans-serif;padding-left:1.1em}
ul.lista li{margin:.35em 0}
a{color:var(--acento)}
.paginacion{display:flex;justify-content:space-between;gap:14px;margin-top:40px;
padding-top:16px;border-top:1px solid var(--linea);font:14px/1.5 system-ui,sans-serif}
footer.sitio{margin-top:48px;padding-top:16px;border-top:1px solid var(--linea);
font:13px/1.6 system-ui,sans-serif;color:var(--suave)}
.art{margin:1.6em 0}
.art .epi{font-weight:700}
table{border-collapse:collapse;width:100%;font:14px/1.5 system-ui,sans-serif;margin:16px 0}
th,td{border:1px solid var(--linea);padding:7px 9px;text-align:left;vertical-align:top}
th{background:var(--caja)}
@media (max-width:600px){.envoltura{padding:18px 16px 56px}h1{font-size:1.4rem}}
"""


def envoltura_html(titulo, descripcion, canonical, cuerpo, jsonld=None):
    ld = ""
    if jsonld:
        ld = '<script type="application/ld+json">%s</script>' % json.dumps(
            jsonld, ensure_ascii=False)
    return """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>%s</title>
<meta name="description" content="%s">
<meta name="robots" content="index,follow,max-snippet:-1,max-image-preview:large">
<link rel="canonical" href="%s">
<meta property="og:title" content="%s">
<meta property="og:description" content="%s">
<meta property="og:type" content="article">
<meta property="og:url" content="%s">
<style>%s</style>
%s
</head>
<body>
<div class="envoltura">
<header class="sitio"><a href="index.html">%s</a></header>
%s
<footer class="sitio">
<p>%s · Repositorio documental de consulta. Los textos se reproducen con fines
informativos; la fuente oficial de cada norma prevalece sobre esta transcripción.</p>
<p><a href="index.html">Índice general</a> · <a href="indice-tematico.html">Índice temático</a></p>
</footer>
</div>
</body>
</html>
""" % (esc(titulo), esc(descripcion[:300]), canonical, esc(titulo),
       esc(descripcion[:300]), canonical, CSS, ld, esc(TITULO_SITIO),
       cuerpo, esc(TITULO_SITIO))


def render_bloques(bloques, doc, articulos_indice, url_pagina):
    partes = []
    usados = set()
    for b in bloques:
        if b["tipo"] == "seccion":
            partes.append("<h3>%s</h3>" % esc(b["texto"]))
        elif b["tipo"] == "articulo":
            anchor = base = "art-" + slugify(b["num"], 24)
            k = 2
            while anchor in usados:
                anchor = "%s-%d" % (base, k)
                k += 1
            usados.add(anchor)
            # El encabezado repite la norma para que el fragmento que devuelva
            # Bing se entienda por sí solo, sin el resto de la página.
            enc = "%s — Artículo %s" % (doc["titulo"], b["num"])
            if b["epigrafe"]:
                enc += ". " + b["epigrafe"]
            partes.append('<h2 id="%s">%s</h2>' % (anchor, esc(enc)))
            if b["texto"]:
                partes.append("<p>%s</p>" % esc(b["texto"]))
            articulos_indice.append({
                "doc": doc["titulo"], "slug": doc["slug"], "num": b["num"],
                "epigrafe": b["epigrafe"], "url": "%s#%s" % (url_pagina, anchor),
            })
        else:
            partes.append("<p>%s</p>" % esc(b["texto"]))
    return "\n".join(partes)


def ficha_html(doc, extra=None):
    filas = [("Tipo", doc["tipo"]), ("Expedido por", doc["emisor"])]
    if doc.get("fecha"):
        filas.append(("Fecha", doc["fecha"]))
    filas.append(("Categoría", doc["categoria"]))
    if extra:
        filas.extend(extra)
    dl = "".join("<dt>%s</dt><dd>%s</dd>" % (esc(k), esc(str(v))) for k, v in filas)
    alias = ""
    if doc.get("alias"):
        alias = "<p><strong>También conocida como:</strong> %s.</p>" % esc(
            "; ".join(doc["alias"]))
    return '<div class="ficha"><dl>%s</dl>%s</div>' % (dl, alias)


# --------------------------------------------------------------------------
# CONSTRUCCIÓN
# --------------------------------------------------------------------------

def construir():
    raiz = os.path.dirname(os.path.abspath(__file__))
    os.chdir(raiz)
    urls = []
    articulos_indice = []
    fichas_docs = []

    for doc in DOCUMENTOS:
        ruta = doc["archivo"]
        if not os.path.exists(ruta):
            print("  ! falta el archivo, se omite:", ruta)
            continue
        print("->", doc["titulo"])
        paginas_pdf = extraer_texto(ruta)
        texto = quitar_encabezados_repetidos(paginas_pdf)
        parrafos = reflujo(texto)
        bloques = estructurar(parrafos)
        grupos = paginar(bloques)

        nombres = []
        for i, g in enumerate(grupos, 1):
            nombres.append("%s-p%02d.html" % (doc["slug"], i) if len(grupos) > 1
                           else "%s-texto.html" % doc["slug"])

        # --- páginas de contenido ---
        for i, (g, nombre) in enumerate(zip(grupos, nombres), 1):
            etiqueta = etiqueta_pagina(g)
            url = "%s/%s" % (SITIO, nombre)
            titulo = "%s — %s | %s" % (doc["titulo"], etiqueta, TITULO_SITIO)
            desc = ("Texto íntegro de %s (%s). %s. Parte %d de %d. %s"
                    % (doc["titulo"], doc["subtitulo"], etiqueta, i, len(grupos),
                       ENTIDAD))
            miga = ('<nav class="miga"><a href="index.html">Índice general</a> › '
                    '<a href="%s.html">%s</a> › %s</nav>'
                    % (doc["slug"], esc(doc["titulo"]), esc(etiqueta)))
            encabezado = ("<h1>%s</h1><p class=\"sub\">%s — parte %d de %d</p>"
                          % (esc(doc["titulo"]), esc(etiqueta), i, len(grupos)))
            cuerpo_txt = render_bloques(g, doc, articulos_indice, nombre)
            nav = ['<div class="paginacion">']
            nav.append('<span>%s</span>' % (
                '<a href="%s">&larr; Parte anterior</a>' % nombres[i - 2]
                if i > 1 else ''))
            nav.append('<span>%s</span>' % (
                '<a href="%s">Parte siguiente &rarr;</a>' % nombres[i]
                if i < len(grupos) else ''))
            nav.append('</div>')
            jsonld = {
                "@context": "https://schema.org",
                "@type": "Legislation" if doc["categoria"] == "Normativa" else "CreativeWork",
                "name": doc["titulo"], "alternateName": doc.get("alias", []),
                "headline": "%s — %s" % (doc["titulo"], etiqueta),
                "inLanguage": "es", "isPartOf": {"@type": "WebSite", "name": TITULO_SITIO,
                                                 "url": SITIO},
                "legislationJurisdiction": "Colombia",
                "publisher": {"@type": "Organization", "name": doc["emisor"]},
                "url": url,
            }
            cuerpo = miga + encabezado + ficha_html(doc) + cuerpo_txt + "".join(nav)
            with open(nombre, "w", encoding="utf-8") as fh:
                fh.write(envoltura_html(titulo, desc, url, cuerpo, jsonld))
            urls.append((url, 0.8))

        # --- portada del documento ---
        portada = "%s.html" % doc["slug"]
        url_p = "%s/%s" % (SITIO, portada)
        lista = "".join(
            '<li><a href="%s">%s</a></li>' % (n, esc(etiqueta_pagina(g)))
            for g, n in zip(grupos, nombres))
        cuerpo = (
            '<nav class="miga"><a href="index.html">Índice general</a> › %s</nav>'
            '<h1>%s</h1><p class="sub">%s</p>%s'
            '<h2>Contenido</h2><ul class="lista">%s</ul>'
            '<p class="aviso">El texto completo está publicado en HTML en las páginas '
            'anteriores. El archivo original se conserva como '
            '<a href="%s">documento fuente</a> para efectos de trazabilidad.</p>'
            % (esc(doc["titulo"]), esc(doc["titulo"]), esc(doc["subtitulo"]),
               ficha_html(doc, [("Partes publicadas", len(grupos))]),
               lista, esc(ruta.replace(" ", "%20"))))
        with open(portada, "w", encoding="utf-8") as fh:
            fh.write(envoltura_html(
                "%s — %s | %s" % (doc["titulo"], doc["subtitulo"], TITULO_SITIO),
                "%s. %s. Texto íntegro consultable en línea. %s"
                % (doc["titulo"], doc["subtitulo"], ENTIDAD),
                url_p, cuerpo))
        urls.append((url_p, 0.9))
        fichas_docs.append((doc, portada, len(grupos)))

    escribir_index(fichas_docs, urls)
    escribir_indice_tematico(articulos_indice, urls)
    escribir_sitemap(urls)
    escribir_robots()
    print("\nArtículos indexados: %d | URLs generadas: %d"
          % (len(articulos_indice), len(urls)))


def escribir_index(fichas, urls):
    secciones = []
    for cat in ("Normativa", "Jurisprudencia", "Plantillas"):
        items = [(d, p, n) for d, p, n in fichas if d["categoria"] == cat]
        if not items:
            continue
        li = "".join(
            '<li><a href="%s"><strong>%s</strong></a> — %s <em>(%d %s)</em></li>'
            % (p, esc(d["titulo"]), esc(d["subtitulo"]), n,
               "parte" if n == 1 else "partes")
            for d, p, n in items)
        secciones.append("<h2>%s</h2><ul class=\"lista\">%s</ul>" % (cat, li))
    cuerpo = (
        "<h1>%s</h1>"
        '<p class="sub">Normativa, jurisprudencia y plantillas jurídicas sobre empleo '
        "público, carrera administrativa y asuntos jurídicos de la %s.</p>"
        '<p>Todos los documentos están publicados con su <strong>texto íntegro en '
        "HTML</strong>, consultable y buscable directamente en esta página. "
        'Ver también el <a href="indice-tematico.html">índice temático por materia</a>.</p>'
        "%s"
        '<p class="aviso">Los textos se reproducen con fines informativos y de consulta '
        "interna. Ante cualquier discrepancia prevalece el texto oficial publicado por "
        "la entidad emisora. Actualizado el %s.</p>"
        % (esc(TITULO_SITIO), esc(ENTIDAD), "".join(secciones), HOY))
    with open("index.html", "w", encoding="utf-8") as fh:
        fh.write(envoltura_html(
            TITULO_SITIO + " — normativa, jurisprudencia y plantillas",
            "Base de conocimiento jurídica de la DIAN: texto íntegro de la Ley 909 de "
            "2004, el Decreto 1083 de 2015, el Decreto Ley 927 de 2023, el CPACA, la "
            "Ley 1712 de 2014, la Constitución Política y jurisprudencia aplicable al "
            "empleo público y la carrera administrativa.",
            SITIO + "/index.html", cuerpo))
    urls.append((SITIO + "/index.html", 1.0))


def escribir_indice_tematico(articulos, urls):
    bloques = []
    for tema, claves in TEMAS.items():
        hits = []
        vistos = set()
        for a in articulos:
            blob = (a["epigrafe"] + " " + a["doc"]).lower()
            if any(k in blob for k in claves):
                clave = (a["doc"], a["num"])
                if clave in vistos:
                    continue
                vistos.add(clave)
                hits.append(a)
        if not hits:
            continue
        filas = "".join(
            "<tr><td>%s</td><td><a href=\"%s\">Artículo %s</a></td><td>%s</td></tr>"
            % (esc(a["doc"]), a["url"], esc(a["num"]), esc(a["epigrafe"][:180]))
            for a in hits[:80])
        bloques.append(
            '<h2 id="%s">%s</h2><table><thead><tr><th>Norma</th><th>Artículo</th>'
            "<th>Materia</th></tr></thead><tbody>%s</tbody></table>"
            % (slugify(tema), esc(tema), filas))
    cuerpo = (
        '<nav class="miga"><a href="index.html">Índice general</a> › Índice temático</nav>'
        "<h1>Índice temático por materia</h1>"
        '<p class="sub">Qué norma y qué artículo consultar según el asunto.</p>'
        + "".join(bloques))
    url = SITIO + "/indice-tematico.html"
    with open("indice-tematico.html", "w", encoding="utf-8") as fh:
        fh.write(envoltura_html(
            "Índice temático por materia | " + TITULO_SITIO,
            "Qué norma y qué artículo aplica según la materia: derecho de petición, "
            "carrera administrativa, traslados y reubicación, situaciones "
            "administrativas, notificaciones y recursos, transparencia y datos "
            "personales.", url, cuerpo))
    urls.append((url, 0.9))


def escribir_sitemap(urls):
    vistos, filas = set(), []
    for u, prio in sorted(urls, key=lambda x: -x[1]):
        if u in vistos:
            continue
        vistos.add(u)
        filas.append("  <url><loc>%s</loc><lastmod>%s</lastmod>"
                     "<changefreq>monthly</changefreq><priority>%.1f</priority></url>"
                     % (html.escape(u), HOY, prio))
    with open("sitemap.xml", "w", encoding="utf-8") as fh:
        fh.write('<?xml version="1.0" encoding="UTF-8"?>\n'
                 '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n'
                 + "\n".join(filas) + "\n</urlset>\n")
    with open("urls.txt", "w", encoding="utf-8") as fh:
        fh.write("\n".join(sorted(vistos)) + "\n")


def escribir_robots():
    with open("robots.txt", "w", encoding="utf-8") as fh:
        fh.write("User-agent: *\nAllow: /\n\n"
                 "User-agent: bingbot\nAllow: /\n\n"
                 "Sitemap: %s/sitemap.xml\n" % SITIO)


if __name__ == "__main__":
    construir()
