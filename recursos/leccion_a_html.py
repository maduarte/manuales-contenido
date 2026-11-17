#!/usr/bin/env python3
"""Convierte el guion de una lección completa en un HTML de lectura, con las respuestas validadas.

    python3 02_Herramientas/leccion_a_html.py U3T3L1.docx -o U3T3L1_vista.html     # .docx bajado del Doc
    python3 02_Herramientas/leccion_a_html.py U3T3L4_guion.md -o U3T3L4_vista.html # guion en .md
    python3 02_Herramientas/leccion_a_html.py tool-result.json -o salida.html      # JSON del conector (docx en base64)

Entradas. El .docx bajado de Google Docs se lee por su estructura de Word (listas, negritas, tablas,
imágenes incrustadas). Un guion en .md se lee directamente (ver "Lectura del .md" más abajo). Si se le
pasa el .docx que produjo guion_a_docx.py desde un .md (sin listas de Word) y el .md está al lado, se
lee el .md, porque ese .docx no conserva la estructura que la vista necesita.

Se pensó para guiones de construcción de conocimiento: varios componentes (Activación, Contexto,
Construcción, Sistematización, Práctica, Ticket), con diálogos, diapositivas, cápsulas, Exploremos,
tablas por completar y árboles. Lo que el estudiante vería; las claves no se muestran hasta comprobar.

Decisiones de diseño (todas tomadas aquí, no en el guion):
  * Componentes. Cada "IPM: Nº componente moodle" abre una sección con su nombre y tipo. Hay una barra
    de navegación fija arriba con los componentes.
  * Preguntas. En el .docx, un párrafo numerado en negrita es el enunciado; los párrafos numerados sin
    negrita que siguen son las alternativas, y se numeran de corrido ("Pregunta n") porque el Doc
    reinicia sus listas. En el .md, la pregunta es la línea con su número escrito ("**4.** …",
    "1. **…**", "**28. …**") y conserva ese número. Alternativas como radio, o casillas si el enunciado
    dice "marca todas", "selecciona todas", "marca las alternativas" o "marca cada".
  * Huecos. "___" es un campo de texto; "___ {a, b}" y "___ [a / b]" son desplegables (la opción en
    negrita es la clave, si la hay). Dentro de una fórmula, \\text{___}, ___ y \\boxed{\\qquad} son campos
    que se dibujan dentro de la fórmula.
  * Claves y validación. Las claves salen de las notas del guion: "IPP: Correcta {b}", "Alternativa
    correcta: {b}", "Respuesta(s) correcta(s): 5; 3; no puede", "Respuestas: a) …; b) …", "Verdadero o
    falso. Respuestas correctas: F; V". Se reparten en orden de lectura entre los campos, desplegables,
    grupos de alternativas y filas V/F de la pregunta. El botón "Comprobar" marca verde o rojo; después
    aparece "Ver respuestas". Si el número de claves no calza con el de respuestas, se avisa por consola.
  * Verdadero/falso. Una tabla con columnas V y F (o "Respuesta" con "___") lleva un radio por columna.
  * Diálogos. Líneas "Nombre: texto" de personajes que hablan al menos dos veces. Cada personaje tiene
    un avatar estilo Playmobil (cabeza amarilla, peinado, polera de color) y el primero va a la
    izquierda, el segundo a la derecha, alternando como en un chat.
  * Diapositivas. Una tabla de una columna con "Diapositiva N" (o precedida por "IPP: cápsula de
    contenido" / "Extracto de noticia" / "Recordemos…") es un carrusel con flechas, puntos y
    teclado. Una cápsula lleva su etiqueta y color.
  * Exploremos / Analicemos. Cuadro plegado, "Exploremos una posible respuesta", que se desbloquea al
    comprobar la pregunta anterior. Lo que el guion pide mostrar "después de validar" queda oculto
    hasta comprobar.
  * Botones "Ver información / Ver tabla". Se muestran como desplegable con lo que sigue.
  * Recursos. Un recurso enlazado al servidor static (.html) se incrusta; uno por producir (IPC,
    "recurso interactivo", "Post recurso") es un recuadro punteado. Las notas de producción (IPP, IPD,
    IPM) no se muestran.
  * Imágenes. Las del Doc se incrustan reducidas; las enlazadas al servidor static se muestran desde
    allí. Las que solo están en Drive o en el Doc ("[imagen: image8.png]") quedan como recuadro.
  * Texto tachado: se ignora (es lo descartado).
  * Fórmulas con KaTeX. Un \\( sin cierre (un error frecuente: \\(27^2 \\cdot 10^4) ) se cierra solo.

Lectura del .md (formato de doc_a_md.py y de los guiones escritos a mano): tablas markdown, "> " de
cita ignorado, $`…`$ como fórmula, \\[ … \\] como fórmula aunque ocupe varias líneas, "IPP: Inicio
(de) diálogo" … "IPP: Fin (de) diálogo", "Cápsula de contenido" (con o sin IPP) hasta "---",
"Diapositiva N" para separar diapositivas, partes "a." / "**b)**", alternativas numeradas "1." / "1)"
o sin numerar (alternativas.py). Lo que viene después de "Propuesta de preguntas para evaluación" no
se muestra.
"""
import argparse
import base64
import html
import io
import json
import os
import re
import sys
import zipfile
import hashlib
import difflib

import docx
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import encabezado_corto  # noqa: E402
from docx.oxml.ns import qn
from PIL import Image

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import alternativas  # noqa: E402

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A_BLIP = "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
R_EMBED = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"

# ------------------------------------------------------------------ lectura del docx


class Doc:
    def __init__(self, ruta):
        if ruta.lower().endswith(".json"):
            j = json.load(open(ruta, encoding="utf-8"))
            b = base64.b64decode(j.get("content") or j.get("fileContent"))
            self.d = docx.Document(io.BytesIO(b))
        else:
            self.d = docx.Document(ruta)
        self.body = self.d.element.body
        self.numfmt = self._numeracion()
        self.imgs = {}

    def _numeracion(self):
        out = {}
        try:
            root = self.d.part.numbering_part.element
        except Exception:
            return out
        abstract = {}
        for an in root.findall(qn("w:abstractNum")):
            aid = an.get(qn("w:abstractNumId"))
            lvl0 = None
            for lvl in an.findall(qn("w:lvl")):
                if lvl.get(qn("w:ilvl")) == "0":
                    nf = lvl.find(qn("w:numFmt"))
                    lvl0 = nf.get(qn("w:val")) if nf is not None else None
            abstract[aid] = lvl0
        for n in root.findall(qn("w:num")):
            aid = n.find(qn("w:abstractNumId")).get(qn("w:val"))
            out[n.get(qn("w:numId"))] = abstract.get(aid)
        return out

    def imagen(self, rid):
        if rid in self.imgs:
            return self.imgs[rid]
        try:
            part = self.d.part.related_parts[rid]
            im = Image.open(io.BytesIO(part.blob))
            im.load()
            if im.width > 1000:
                im = im.resize((1000, int(im.height * 1000 / im.width)))
            buf = io.BytesIO()
            if im.mode in ("RGBA", "LA", "P"):
                im.convert("RGBA").save(buf, "PNG", optimize=True)
                mime = "image/png"
            else:
                im.convert("RGB").save(buf, "JPEG", quality=82)
                mime = "image/jpeg"
            uri = "data:%s;base64,%s" % (mime, base64.b64encode(buf.getvalue()).decode())
            vacia = im.width * im.height < 100
        except Exception:
            uri, vacia = None, True
        self.imgs[rid] = None if vacia else uri
        return self.imgs[rid]


def runs_de(p):
    """Devuelve lista de (texto, negrita, es_imagen_rid) de un párrafo, sin lo tachado."""
    out = []
    for r in p.iter(qn("w:r")):
        if any(a.tag in (W + "drawing", W + "pict", W + "txbxContent") or a.tag.endswith("}AlternateContent")
               for a in r.iterancestors()):
            continue
        rpr = r.find(qn("w:rPr"))
        neg = tach = False
        if rpr is not None:
            b = rpr.find(qn("w:b"))
            neg = b is not None and b.get(qn("w:val")) not in ("0", "false")
            s = rpr.find(qn("w:strike"))
            tach = s is not None and s.get(qn("w:val")) not in ("0", "false")
            s2 = rpr.find(qn("w:dstrike"))
            tach = tach or (s2 is not None)
        if tach:
            continue
        for ch in r:
            if ch.tag == W + "t":
                out.append((ch.text or "", neg, None))
            elif ch.tag == W + "br":
                out.append(("\n", False, None))
            elif ch.tag == W + "tab":
                out.append((" ", False, None))
            elif ch.tag == W + "drawing":
                for bl in ch.iter(A_BLIP):
                    out.append(("", False, bl.get(R_EMBED)))
    return out


def para(doc, p):
    """Párrafo -> dict con texto marcado (**negrita**), texto plano, imágenes, lista."""
    items = runs_de(p)
    texto, plano, imgs = "", "", []
    neg_chars = tot_chars = 0
    for t, neg, rid in items:
        if rid:
            u = doc.imagen(rid)
            if u:
                imgs.append(u)
            continue
        plano += t
        if t.strip():
            tot_chars += len(t.strip())
            if neg:
                neg_chars += len(t.strip())
        if neg and t.strip():
            texto += "**" + t + "**"
        else:
            texto += t
    texto = re.sub(r"\*\*\s*\*\*", " ", texto)
    texto = re.sub(r"\*\*(\s+)\*\*", r"\1", texto)
    num = None
    ppr = p.find(qn("w:pPr"))
    if ppr is not None:
        npr = ppr.find(qn("w:numPr"))
        if npr is not None and npr.find(qn("w:numId")) is not None:
            nid = npr.find(qn("w:numId")).get(qn("w:val"))
            fmt = doc.numfmt.get(nid)
            if nid != "0":
                num = dict(id=nid, fmt=fmt)
    neg = neg_chars / tot_chars if tot_chars else 0
    if num is None:
        # .docx de guion_a_docx.py: la numeración va escrita ("11." en negrita para el enunciado,
        # "a." sin negrita para las alternativas), no como lista de Word.
        primero = next((n for t, n, rid in items if not rid and t.strip()), False)
        m = re.match(r"^\s*(\d+)\.\s+", plano)
        if m and primero:
            num, neg = dict(id="texto", fmt="decimal"), 1.0
        else:
            m = re.match(r"^\s*([a-h])\.\s+", plano)
            if m and not primero:
                num, neg = dict(id="texto", fmt="lowerLetter"), 0
        if num is not None:
            plano = plano[m.end():]
            texto = re.sub(r"^\s*(\*\*)?\s*\w+\.\s*(\*\*)?\s*", "", texto, count=1)
    return dict(texto=texto, plano=plano, imgs=imgs, neg=neg, num=num)


def celda(doc, tc):
    """Celda -> lista de bloques: ('p', dict) o ('t', filas)."""
    out = []
    for el in tc.iterchildren():
        if el.tag == W + "p":
            d = para(doc, el)
            if d["plano"].strip() or d["imgs"]:
                out.append(("p", d))
        elif el.tag == W + "tbl":
            out.append(("t", tabla(doc, el)))
    return out


def tabla(doc, tbl):
    filas = []
    for tr in tbl.findall(qn("w:tr")):
        filas.append([celda(doc, tc) for tc in tr.findall(qn("w:tc"))])
    return filas


RUIDO = re.compile(r"^(IP[PDCM]:.*|\S+(?: \S+)? ?\d?\.(jpg|jpeg|png)|[.,;]+)$", re.I)


def lineas_de(d, filtrar=False):
    """Divide un párrafo en líneas conservando la negrita. Con filtrar, quita notas IPP/IPD y nombres de archivo."""
    res = []
    for l, lp in zip(d["texto"].split("\n"), d["plano"].split("\n")):
        if lp.strip():
            if filtrar and RUIDO.match(lp.strip()):
                continue
            res.append((l.strip(), lp.strip()))
    return res


# ------------------------------------------------------------------ texto -> html


def arreglar_mate(t):
    """Cierra \\( sin \\) dentro de una misma línea."""
    def fix(m):
        s = m.group(0)
        if "\\(" in s and s.count("\\(") > s.count("\\)"):
            i = s.rfind("\\(")
            j = s.find(")", i + 2)
            while j != -1 and s[i + 2:j].count("(") != s[i + 2:j].count(")"):
                j = s.find(")", j + 1)
            if j != -1:
                return s[:j] + "\\)" + s[j + 1:]
            return s + "\\)"
        return s
    return re.sub(r"[^\n]+", fix, t)


# ------------------------------------------------------------------ huecos, controles y claves
#
# Cada respuesta que el estudiante da (campo, desplegable, grupo de alternativas, fila V/F) es un
# "control". Al escribirlo se registra en _CTL y deja en el HTML la marca data-k='\x01N\x01'. Cuando se
# arma la tarjeta de la pregunta, envolver_pregunta() reparte las claves del guion entre esas marcas,
# en orden de lectura, y las reemplaza por la clave en base64. Las marcas que quedan sin clave se
# borran al final (generar).

_CTL = {}
_hz = [0]


def latex_a_texto(s):
    """Fórmula sencilla -> texto plano con símbolos (para opciones de desplegable y claves)."""
    s = s.replace("\\(", "").replace("\\)", "").replace("$", "").replace("`", "")
    s = re.sub(r"\\d?frac\{([^{}]*)\}\{([^{}]*)\}", r"\1/\2", s)
    s = re.sub(r"\\(?:text|mathrm|textbf)\{([^{}]*)\}", r"\1", s)
    for a, b in (("\\div", "÷"), ("\\cdot", "·"), ("\\times", "×"), ("\\leq", "≤"), ("\\geq", "≥"),
                 ("\\le", "≤"), ("\\ge", "≥"), ("\\{", "{"), ("\\}", "}"), ("\\,", ""), ("\\ ", " "),
                 ("\\!", ""), ("\;", " ")):
        s = s.replace(a, b)
    return re.sub(r"\s+", " ", s).strip()


def nrm(s):
    """Normaliza una respuesta para compararla (la misma regla que usa el JS de la página)."""
    import unicodedata
    s = unicodedata.normalize("NFD", latex_a_texto(s or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn").lower()
    s = re.sub(r"[·×*]", "·", s)
    s = re.sub(r"[\s{}()\[\],\\]", "", s)
    return re.sub(r"[.;:]+$", "", s)


def b64(s):
    return base64.b64encode(s.encode("utf-8")).decode()


def nuevo_control(tipo, **datos):
    n = len(_CTL) + 1
    _CTL[n] = dict(tipo=tipo, **datos)
    return n, " data-k='\x01%d\x01' data-tipo='%s'" % (n, tipo)


def control_input(ancho=None):
    _, atr = nuevo_control("input")
    est = " style='width:%dch'" % ancho if ancho else ""
    return "<input type='text' class='hz'%s%s autocomplete='off' spellcheck='false' aria-label='Respuesta'>" % (atr, est)


def opciones_de(cuerpo):
    """'a / b' , 'a; b' o 'a, b' -> [(texto visible, es_clave_en_negrita)]."""
    if re.search(r"\s/|/\s", cuerpo):
        partes = re.split(r"\s+/\s*|\s*/\s+", cuerpo)
    elif ";" in cuerpo:
        partes = re.split(r";", cuerpo)
    else:
        partes = cuerpo.split(",")
    out = []
    for p in partes:
        neg = "**" in p
        t = latex_a_texto(p.replace("**", "")).strip(" /")
        if t:
            out.append((t, neg))
    return out


def control_select(opciones):
    kb = next((t for t, neg in opciones if neg), None)
    if len([1 for _, neg in opciones if neg]) != 1:
        kb = None
    _, atr = nuevo_control("select", opciones=[t for t, _ in opciones], kb=kb)
    return "<select class='hz'%s><option value=''>Elige…</option>%s</select>" % (
        atr, "".join("<option>%s</option>" % html.escape(t) for t, _ in opciones))


# "___ {a, b}" o "___ [a / b]" (en el .md los corchetes vienen escapados: \[a / b\])
RE_SEL = re.compile(r"(?:\*\*)?_{3,}(?:\*\*)?(?:[ \t]|\\ )*(?:\\\[([^\[\]]+?)\\\]|\[([^\[\]]+?)\]|(?<!\\)\{([^{}]+)\})")
RE_INPUT = re.compile(r"(?:\*\*)?_{3,}(?:\*\*)?")
# huecos dentro de una fórmula
RE_HZ_MATE = re.compile(r"\\text\{\s*_{2,}\s*\}|\\boxed\{\s*(?:\\q?quad\s*)+\}|_{2,}\s*\\?\{([^{}]*,[^{}]*)\\?\}|_{3,}"
                        r"|\x02(\d+)\x02")


def hueco_en_mate(m, pendientes, selects):
    """Reemplaza un hueco dentro de la fórmula por una caja con id; el control queda 'pendiente' y el
    JS lo mueve adentro de la caja cuando KaTeX dibujó la fórmula."""
    _hz[0] += 1
    hid = "hz-%d" % _hz[0]
    if m.group(1):                                     # "__{1,2,**3**,4}": desplegable dentro de la fórmula
        ctl = control_select(opciones_de(m.group(1)))
    elif m.group(2):
        ctl = selects[int(m.group(2))]
    else:
        ctl = control_input()
    pendientes.append("<span class='pend' data-hz='%s'>%s</span>" % (hid, ctl))
    return "\\htmlId{%s}{\\boxed{\\phantom{0000}}}" % hid


def texto_con_huecos(p, selects):
    """Texto sin fórmulas -> HTML con negritas, campos y desplegables."""
    out, pos = [], 0
    for m in re.finditer(r"\x02(\d+)\x02|\\text\{\s*_{2,}\s*\}|(?:\*\*)?_{3,}(?:\*\*)?", p):
        out.append(p[pos:m.start()])
        out.append("\x03%d\x03" % len(_PEND_TXT))
        _PEND_TXT.append(selects[int(m.group(1))] if m.group(1) else control_input())
        pos = m.end()
    out.append(p[pos:])
    t = html.escape("".join(out), quote=False)
    t = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", t)
    t = t.replace("*", "")
    t = re.sub(r"\\[\]\)]", "", t)                   # cierres de fórmula sueltos
    return re.sub(r"\x03(\d+)\x03", lambda m: _PEND_TXT[int(m.group(1))], t)


_PEND_TXT = []


def inline(t, huecos=True):
    t = re.sub(r"\\(\*\*)([()\[\]])", r"\1\\\2", t)     # negrita que parte el delimitador: "\**(" -> "**\("
    t = arreglar_mate(t)
    t = re.sub(r"(\\\])\s*[.,]\s*$", r"\1", t)
    if huecos:
        t = re.sub(r"(_{3,})\s*\*\*\s*\{", r"\1 {", t)        # "___ **{**a**, b}**" -> "___ {**a**, b}"
        t = re.sub(r"(_{3,} \{[^{}]*\})\*\*", r"\1", t)
    selects = []
    if huecos:
        def sel(m):
            selects.append(control_select(opciones_de(m.group(1) or m.group(2) or m.group(3))))
            return "\x02%d\x02" % (len(selects) - 1)
        t = RE_SEL.sub(sel, t)
    partes = re.split(r"(\\\(.*?\\\)|\\\[.*?\\\])", t, flags=re.S)
    out = []
    for p in partes:
        if p.startswith("\\(") or p.startswith("\\["):
            pend = []
            if huecos:
                p = RE_HZ_MATE.sub(lambda m: hueco_en_mate(m, pend, selects), p)
            out.append(html.escape(p, quote=False) + "".join(pend))
        elif out and out[-1].startswith("\\[") and re.fullmatch(r"[\s.,;]*(\\[\])])?[\s.,;]*", p):
            continue                                       # puntuación suelta tras una fórmula de bloque
        elif huecos:
            out.append(texto_con_huecos(p, selects))
        else:
            p = html.escape(p, quote=False)
            p = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", p)
            out.append(p.replace("*", ""))
    return "".join(out)


def tokens_clave(linea):
    """Línea de clave del guion -> lista de respuestas en orden, o None si la línea no es una clave.

    Reconoce "IPP: Correcta {b}", "Alternativa correcta: {b}", "Respuesta correcta: b).",
    "Respuestas correctas, en orden: 5; 3; no puede", "Respuestas: a) 16; b) 6", "respuestas {8, 3, No}"
    y "Correcta b." (sin IPP). Corta las instrucciones que siguen ("No completar…", "Mostrar…").
    """
    s = re.sub(r"^[>\s*]+|[\s*]+$", "", linea)
    s = re.sub(r"^IP[PMDC]\s*:\s*", "", s)
    if re.search(r"(?i)\bsin respuesta correcta\b", s):
        return None
    m = re.search(r"(?i)\b(?:correctas?|respuestas?)\b[^{}]{0,40}?((?:\{[^{}]*\}[\s.]*)+)$", s)
    if m:
        return tokens_llaves(re.findall(r"\{([^{}]*)\}", m.group(1)))
    m = re.search(r"(?i)\b(?:correctas?|respuestas?)\b(?:\s+correctas?)?(?:,?\s*en orden(?: de lectura)?)?\s*:\s*(.*)$", s)
    if not m:
        m = re.match(r"(?i)^correctas?\s+([a-e]\b.*)$", s)
    if not m:
        return None
    resto = m.group(1).strip()
    if not resto:
        return []
    resto = re.split(r"\.\s+(?=[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+\s+\S+\s+\S+)", resto)[0]
    resto = re.sub(r"(?i)\s*\ben la retroalimentaci.*$", "", resto)
    # partes "a) …" / "a. …": cada respuesta queda marcada con su parte ("\x05a\x05respuesta")
    trozos = re.split(r"(?:^|(?<=[\s;.]))([a-e])\)\s*|(?:^|(?<=[\s;]))([a-e])\.\s+(?=\S)", resto)
    toks, parte = [], ""
    for k, x in enumerate(trozos):
        if k % 3 == 1:
            parte = x or trozos[k + 1] or ""
            continue
        if k % 3 == 2 or x is None:
            continue
        for y in x.split(";"):
            y = re.sub(r"(?i)^\s*alternativa\s+", "", y).strip()
            y = re.sub(r"[.\s]+$", "", y)
            if not y:
                continue
            pre = "\x05%s\x05" % parte if parte else ""
            doble = re.match(r"(?i)^(.*?)\s+en (?:ambos|los dos) (?:espacios|campos)$", y)
            if doble:
                toks += [pre + latex_a_texto(doble.group(1))] * 2
            else:
                toks.append(pre + latex_a_texto(y))
    return toks


def tokens_llaves(grupos):
    """Contenido de "{…}" en una clave: "{b}", "{a, c}", "{8, 3, No}", "{una · tres · 3}", "{2.a: b · 2.b: c}",
    "{a: 2 · 1} {b: 2 · 4}", "{fila 3: Color · fila 4: Claro}"."""
    toks, parte = [], ""
    for g in grupos:
        g = g.replace("**", "").strip()
        if re.fullmatch(r"[a-eA-E](\s*(,|;|y)\s*[a-eA-E])*", g):
            toks.append(g)
            continue
        piezas = g.split("·") if "·" in g else g.split(",")
        for x in piezas:
            x = x.strip()
            m = re.match(r"^(?:\d+\.)?([a-e])\s*:\s*(.*)$", x)
            if m:
                parte, x = m.group(1), m.group(2)
            x = re.sub(r"(?i)^(fila|etapa|parte|casilla|campo)\s*\d+\s*:\s*", "", x).strip()
            sub = [y.strip() for y in x.split(",")] if "·" in g else [x]
            for y in sub:
                if y:
                    toks.append(("\x05%s\x05" % parte if parte else "") + latex_a_texto(y))
    return toks


def resolver_clave(ctl, tok):
    """Ajusta una clave al control que la recibe (letra para alternativas, opción exacta para desplegables)."""
    tipo = ctl["tipo"]
    if tok is None:
        return ctl.get("kb")
    if tipo == "select":
        ops = ctl["opciones"]
        for o in ops:
            if nrm(o) == nrm(tok):
                return o
        empieza = [o for o in ops if nrm(tok) and nrm(o).startswith(nrm(tok))]
        if len(empieza) == 1:
            return empieza[0]
        mejor = max(ops, key=lambda o: difflib.SequenceMatcher(None, nrm(o), nrm(tok)).ratio(), default=None)
        if mejor and difflib.SequenceMatcher(None, nrm(mejor), nrm(tok)).ratio() >= 0.6:
            return mejor
        return ctl.get("kb")
    if tipo in ("radio", "multi"):
        letras = re.findall(r"(?i)(?:^|[\s,;y])([a-e])(?=$|[\s,;.)])", " " + tok)
        if tipo == "radio" and re.match(r"(?i)^\s*[a-e]\s*(?:[.)]|$)", tok):
            return tok.strip()[0].lower()
        if tipo == "multi":
            nums = re.findall(r"\d+", tok)
            if nums and ctl.get("filas"):
                return ",".join(nums)
            if letras:
                return ",".join(x.lower() for x in letras)
        ops = ctl.get("opciones") or []
        for j, o in enumerate(ops):
            if nrm(o) == nrm(tok):
                return "abcdefgh"[j]
        dentro = [j for j, o in enumerate(ops) if nrm(tok) and nrm(tok) in nrm(o)]
        if len(dentro) == 1:
            return "abcdefgh"[dentro[0]]
        return None
    if tipo == "vf":
        return tok.strip()[:1].upper() if re.match(r"(?i)^[vf]", tok.strip()) else None
    return tok


def calza(ctl, tok):
    """¿La clave puede ser la de este control? (para repartir cuando el número no coincide)"""
    t = ctl["tipo"]
    if t == "radio":
        return resolver_clave(ctl, tok) is not None
    if t == "vf":
        return re.match(r"(?i)^[vf]$", tok.strip()) is not None
    if t == "select":
        return any(nrm(o) == nrm(tok) or (nrm(tok) and nrm(o).startswith(nrm(tok))) or
                   difflib.SequenceMatcher(None, nrm(o), nrm(tok)).ratio() >= 0.6 for o in ctl["opciones"])
    if t == "multi":
        return len(re.findall(r"\d+|\b[a-e]\b", tok)) > 1
    return re.match(r"(?i)^\s*[a-e]\s*$", tok) is None


def repartir(ids, claves):
    """Reparte las claves en orden, saltando las que no calzan con el tipo de respuesta."""
    asign, j = {}, 0
    for i in ids:
        for k in range(j, len(claves)):
            if calza(_CTL[i], claves[k]):
                asign[i] = claves[k]
                j = k + 1
                break
    return asign


def envolver_pregunta(n, enun_html, cuerpo, claves, etiqueta="Pregunta"):
    """Tarjeta de pregunta: reparte las claves entre los controles y agrega el botón Comprobar."""
    ids, partes, parte = [], {}, ""
    for m in re.finditer(r"\x01(\d+)\x01|\x05([a-h])\x05", cuerpo):
        if m.group(2):
            parte = m.group(2)
        else:
            ids.append(int(m.group(1)))
            partes[int(m.group(1))] = parte
    cuerpo = re.sub(r"\x05[a-h]\x05", "", cuerpo)
    etiq = [re.match(r"^(?:\x05([a-h])\x05)?(.*)$", c, re.S).groups() for c in (claves or [])]
    claves = [t for _, t in etiq]
    con_kb = [i for i in ids if _CTL[i].get("kb")]
    sin_kb = [i for i in ids if i not in con_kb]

    def todas_calzan(pares):
        return all(calza(_CTL[i], t) for i, t in pares)
    if len(claves) == len(ids) and todas_calzan(zip(ids, claves)):
        asign = dict(zip(ids, claves))
    elif claves and len(claves) == len(sin_kb) and todas_calzan(zip(sin_kb, claves)):
        asign = dict(zip(sin_kb, claves))
    else:
        asign = {}
        if any(p for p, _ in etiq) and any(partes.values()):
            for letra in sorted(set(p for p, _ in etiq if p)):
                asign.update(repartir([i for i in ids if partes[i] == letra], [t for p, t in etiq if p == letra]))
        else:
            asign = repartir(ids, claves)
        if claves or (ids and not con_kb):
            print("[claves] %s %s: %d claves para %d respuestas%s" % (
                etiqueta, n, len(claves), len(ids), ("; se reparten por parte y tipo" if claves else "")), file=sys.stderr)
    nclave = 0
    for i in ids:
        k = resolver_clave(_CTL[i], asign.get(i))
        if k:
            nclave += 1
            cuerpo = cuerpo.replace("\x01%d\x01" % i, b64(k), 1)
    abierta = "<textarea" in cuerpo
    if nclave:
        boton = "Comprobar"
    elif ids or abierta:
        boton = "Enviar"
    else:
        boton = "Continuar"
    return ("<div class='card q' data-n='%s'><span class='qn'>%s %s</span>%s%s"
            "<div class='acc'><button type='button' class='chk'>%s</button><span class='res' aria-live='polite'></span>"
            "<button type='button' class='versol' hidden>Ver respuestas</button></div></div>") % (
        html.escape(str(n)), etiqueta, html.escape(str(n)), enun_html, cuerpo, boton)


def ul_alternativas(textos, multi, g, imagenes=False):
    tipo = "checkbox" if multi else "radio"
    _, atr = nuevo_control("multi" if multi else "radio", opciones=[latex_a_texto(t) for t in textos] if not imagenes else [])
    if imagenes:
        lis = "".join("<li data-i='%s'><label><input type='%s' name='%s'><span class='lab'>Opción %s</span>%s</label></li>" % (
            "abcdefgh"[j], tipo, g, "abcdefgh"[j], render_img(u)) for j, u in enumerate(textos))
        return "<ul class='opts img'%s>%s</ul>" % (atr, lis)
    lis = "".join("<li data-i='%s'><label><input type='%s' name='%s'><span>%s</span></label></li>" % (
        "abcdefgh"[j], tipo, g, inline(t)) for j, t in enumerate(textos))
    return "<ul class='opts'%s>%s</ul>" % (atr, lis)


# ------------------------------------------------------------------ avatares

PALETAS = [
    ("#d9468a", "#4a2c17", "ponytail"),   # polera rosa, pelo castaño, cola
    ("#2f80ed", "#1f1410", "short"),      # polera azul, pelo oscuro, corto
    ("#27ae60", "#7a4a1d", "long"),
    ("#f2994a", "#202020", "cap"),
    ("#9b51e0", "#b5651d", "short"),
    ("#eb5757", "#3a2616", "ponytail"),
]
CONOCIDOS = {"inés": 0, "ines": 0, "seba": 1}
# Rótulos que llevan dos puntos pero no son personajes ("Respuesta: ___ diseños.").
NO_PERSONAJES = {"Respuesta", "Respuestas", "Nota", "Ejemplo", "Clave", "Tabla", "Imagen", "Pregunta",
                 "Caso", "Total", "Recurso", "Link", "Fuente", "Objetivo", "Contexto", "Alternativa"}


def avatar_svg(nombre, idx):
    k = CONOCIDOS.get(nombre.lower())
    if k is None:
        # peinado según el nombre (los terminados en "a" con cola o pelo largo) y color según el orden
        grupo = [0, 2, 5] if nombre.lower().rstrip("s").endswith("a") else [1, 3, 4]
        k = grupo[(idx + int(hashlib.md5(nombre.encode()).hexdigest(), 16)) % len(grupo)]
    polera, pelo, peinado = PALETAS[k]
    piel = "#f6cf57"
    s = ['<svg class="av" viewBox="0 0 100 124" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="%s">' % html.escape(nombre)]
    # cuerpo estilo Playmobil: torso trapecio, brazos con manos en C
    s.append('<rect x="30" y="72" width="40" height="40" rx="9" fill="%s"/>' % polera)
    s.append('<rect x="14" y="76" width="16" height="30" rx="8" fill="%s"/>' % polera)
    s.append('<rect x="70" y="76" width="16" height="30" rx="8" fill="%s"/>' % polera)
    s.append('<circle cx="22" cy="110" r="7" fill="%s"/><circle cx="78" cy="110" r="7" fill="%s"/>' % (piel, piel))
    s.append('<rect x="31" y="110" width="17" height="14" rx="3" fill="#3b3b4f"/><rect x="52" y="110" width="17" height="14" rx="3" fill="#3b3b4f"/>')
    s.append('<rect x="43" y="62" width="14" height="14" rx="3" fill="%s"/>' % piel)
    # peinado de atrás
    if peinado == "long":
        s.append('<path d="M20 40 Q20 8 50 8 Q80 8 80 40 L82 78 Q70 84 66 70 L34 70 Q30 84 18 78 Z" fill="%s"/>' % pelo)
    if peinado == "ponytail":
        s.append('<ellipse cx="82" cy="52" rx="9" ry="16" fill="%s" transform="rotate(-18 82 52)"/>' % pelo)
    # cabeza
    s.append('<ellipse cx="50" cy="40" rx="27" ry="28" fill="%s" stroke="#d9a92b" stroke-width="1.5"/>' % piel)
    s.append('<ellipse cx="23" cy="42" rx="4" ry="6" fill="%s" stroke="#d9a92b" stroke-width="1"/><ellipse cx="77" cy="42" rx="4" ry="6" fill="%s" stroke="#d9a92b" stroke-width="1"/>' % (piel, piel))
    # pelo de arriba
    if peinado == "short":
        s.append('<path d="M23 38 Q22 10 50 10 Q78 10 77 38 Q70 26 58 24 Q44 30 30 24 Q26 28 23 38 Z" fill="%s"/>' % pelo)
    elif peinado == "cap":
        s.append('<path d="M22 36 Q24 8 50 8 Q76 8 78 36 Z" fill="%s"/><rect x="40" y="30" width="46" height="8" rx="4" fill="%s"/>' % (pelo, pelo))
    else:
        s.append('<path d="M23 40 Q21 8 50 8 Q79 8 77 40 Q66 22 50 24 Q34 22 23 40 Z" fill="%s"/>' % pelo)
    # cara
    s.append('<circle cx="40" cy="42" r="3.2" fill="#222"/><circle cx="60" cy="42" r="3.2" fill="#222"/>')
    s.append('<path d="M33 35 Q40 32 46 35 M54 35 Q60 32 67 35" stroke="#5a3b1a" stroke-width="2" fill="none" stroke-linecap="round"/>')
    s.append('<path d="M38 53 Q50 63 62 53" stroke="#7a3b1a" stroke-width="2.6" fill="none" stroke-linecap="round"/>')
    s.append("</svg>")
    return "".join(s)


# ------------------------------------------------------------------ modelo


class Comp:
    def __init__(self, n):
        self.n = n
        self.titulo = ""
        self.tipo = ""
        self.items = []


def es_vf(filas):
    if not filas:
        return False
    cab = [" ".join(b[1]["plano"] for b in c if b[0] == "p").strip() for c in filas[0]]
    return len(cab) >= 3 and cab[-2:] == ["V", "F"]


def texto_celda(c):
    return " ".join(b[1]["plano"].strip() for b in c if b[0] == "p").strip()


def parsear(doc):
    # 1) personajes: nombres que hablan al menos 2 veces
    todo = [p for p in doc.body.iter(qn("w:p"))]
    cuenta = {}
    for p in todo:
        d = para(doc, p)
        for _, lp in lineas_de(d):
            m = re.match(r"^([A-ZÁÉÍÓÚÑ][a-záéíóúñ]{1,12}):\s", lp)
            if m:
                cuenta[m.group(1)] = cuenta.get(m.group(1), 0) + 1
    personajes = [n for n, c in cuenta.items() if c >= 2 and n not in NO_PERSONAJES]

    meta = {}
    comps = []
    cur = None
    q = None
    nq = [0]
    bandera = dict(capsula=None, saltar_tabla=False, boton=None, objetivo=False, pend_nota=None)

    def cerrar_q():
        nonlocal q
        if q is not None:
            cur.items.append(q)
            q = None

    def agregar(b, dentro_q=True):
        if cur is None:
            return
        if q is not None and dentro_q and not q["opciones"]:
            q["bloques"].append(b)
        else:
            cerrar_q()
            cur.items.append(b)

    def lineas_dialogo(lp, lt):
        m = re.match(r"^([A-ZÁÉÍÓÚÑ][a-záéíóúñ]{1,12}):\s*(.*)$", lp)
        if m and m.group(1) in personajes:
            resto = re.sub(r"^\*{0,2}[^:]{1,14}:\*{0,2}\s*", "", lt)
            return m.group(1), resto.replace("**", "").strip()
        return None

    for el in doc.body.iterchildren():
        if el.tag == W + "tbl":
            filas = tabla(doc, el)
            if cur is None:
                meta.setdefault("tablas", []).append(filas)
                continue
            if bandera["saltar_tabla"]:
                bandera["saltar_tabla"] = False
                continue
            clasificar_tabla(filas, bandera, agregar, cerrar_q, lambda: q, doc)
            continue
        if el.tag != W + "p":
            continue
        d = para(doc, el)
        for u in d["imgs"]:
            agregar(dict(k="img", uri=u))
        lins = lineas_de(d)
        if d["num"] and not d["imgs"] and lins:
            # párrafo de lista: enunciado, alternativa o viñeta
            lt, lp = lins[0]
            lt = " ".join(x[0] for x in lins)
            lp = " ".join(x[1] for x in lins)
            if cur is None:
                continue
            if d["num"]["fmt"] == "bullet":
                b = dict(k="ul", items=[lt])
                if q is not None and not q["opciones"] and q["bloques"] and q["bloques"][-1]["k"] == "ul":
                    q["bloques"][-1]["items"].append(lt)
                elif q is None and cur.items and cur.items[-1].get("k") == "ul":
                    cur.items[-1]["items"].append(lt)
                else:
                    agregar(b)
                continue
            if d["neg"] >= 0.6:
                cerrar_q()
                if "Post recurso" in lp:
                    continue
                nq[0] += 1
                q = dict(k="q", n=nq[0], stem=lt, bloques=[], opciones=[], claves=[],
                         multi=bool(re.search(r"[Mm]arca (todas|las alternativas|cada)\b(?!.*verdadera)", lp)) and "(V)" not in lp)
                bandera["ultima_q"] = q
                bandera["pend_nota"] = None
                continue
            if q is not None:
                q["opciones"].append(lt)
                continue
            agregar(dict(k="p", t=lt))
            continue
        for lt, lp in lins:
            low = lp.lower()
            if low.startswith("ipm:"):
                cerrar_q()
                n = re.search(r"(\d+)", lp)
                cur = Comp(int(n.group(1)) if n else len(comps) + 1)
                bandera["ultima_q"] = None
                comps.append(cur)
                continue
            if cur is not None and low.startswith("nombre del componente"):
                cur.titulo = re.sub(r"^.*?[:：]\s*[“\"]?", "", lp).strip("”\" ")
                continue
            if cur is not None and low.startswith("tipo de componente"):
                cur.tipo = re.sub(r"^.*?:\s*", "", lp).strip(" .")
                continue
            if cur is None:
                continue
            if set(lp) <= set("_- ") or low.startswith("ipp: inicio elemento") or low.startswith("ipp: fin elemento"):
                if low.startswith("ipp: fin elemento") or low.startswith("ipp: inicio elemento"):
                    pass
                continue
            if low.startswith("ipp: cápsula de objetivo") or low.startswith("ipp: capsula de objetivo"):
                bandera["objetivo"] = True
                continue
            if low.startswith("ipp: cápsula de contenido"):
                bandera["capsula"] = "Cápsula de contenido"
                continue
            if low.startswith("recordemos"):
                bandera["recordemos"] = True
                agregar(dict(k="titulo", t="Recordemos…"))
                continue
            if low == "cápsula de contenido":
                continue
            if re.match(r"^\S+(?: \S+)? ?\d?\.(jpg|jpeg|png)$", low):
                continue
            if low.startswith("ipp: extracto de noticia"):
                bandera["capsula"] = "Extracto de noticia"
                continue
            if low.startswith("ipp:") or low.startswith("ipd:"):
                toks = tokens_clave(lp) if low.startswith("ipp:") else None
                if toks and bandera.get("ultima_q") is not None:
                    bandera["ultima_q"]["claves"].extend(toks)
                if re.match(r"ipp:\s*correcta\s*\.?$", low):
                    bandera["saltar_tabla"] = True
                m = re.match(r"ipp:\s*bot[óo]n\s*[“\"]?(.+?)[”\"]?\s*\.?$", lp, re.I)
                if m:
                    bandera["boton"] = m.group(1).strip()
                continue
            if low.startswith("ipc:"):
                if re.search(r"recurso|diagrama|imagen", low) and "ideas" not in low:
                    agregar(dict(k="recurso", t=re.sub(r"^IPC:\s*", "", lp)))
                continue
            if low.startswith("acá hay que crear un recurso"):
                bandera["pend_nota"] = "recurso interactivo"
                agregar(dict(k="recurso", t="Recurso interactivo: árbol parcial y fila de casillas"))
                continue
            if bandera["pend_nota"] and (lp.startswith("*") or lp.startswith("-")):
                continue
            dd = lineas_dialogo(lp, lt)
            if dd:
                cerrar_q()
                if cur.items and cur.items[-1].get("k") == "dialogo":
                    cur.items[-1]["lineas"].append(dd)
                else:
                    cur.items.append(dict(k="dialogo", lineas=[dd]))
                continue
            if bandera["objetivo"]:
                cerrar_q()
                if cur.items and cur.items[-1].get("k") == "objetivo":
                    cur.items[-1]["lineas"].append(lt)
                else:
                    cur.items.append(dict(k="objetivo", lineas=[lt]))
                continue
            if bandera["boton"] is not None:
                pass
            if q is not None and q["opciones"]:
                cerrar_q()
            agregar(dict(k="p", t=lt))
    cerrar_q()
    return meta, comps, personajes


def clasificar_tabla(filas, bandera, agregar, cerrar_q, get_q, doc):
    """Decide qué es una tabla y la agrega."""
    ncol = max(len(f) for f in filas) if filas else 0
    todas = [c for f in filas for c in f]
    textos = [texto_celda(c) for c in todas]
    # botón: lo que sigue va dentro de un desplegable
    boton = bandera["boton"]
    bandera["boton"] = None

    def poner(b):
        if boton:
            agregar(dict(k="boton", texto=boton, contenido=[b]))
        else:
            agregar(b)

    # carrusel
    es_diap = any(re.match(r"^\W*Diapositiva\s*\d+", t) for t in textos)
    if ncol == 1 and (es_diap or bandera["capsula"]):
        etiqueta = bandera["capsula"] or ("Recordemos" if bandera.get("recordemos") else "Diapositivas")
        bandera["capsula"] = None
        bandera["recordemos"] = False
        cerrar_q()
        agregar(dict(k="carrusel", etiqueta=etiqueta, filas=[f[0] for f in filas]), dentro_q=False)
        return
    bandera["capsula"] = None
    if ncol == 1 and len(filas) == 1:
        t = textos[0]
        if re.match(r"^\W*(Exploremos|Analicemos)", t):
            cerrar_q()
            agregar(dict(k="explore", celda=todas[0]), dentro_q=False)
            return
        if re.match(r"^\W*Situaci[oó]n", t) or len(t) > 60:
            cerrar_q()
            agregar(dict(k="caja", celda=todas[0]), dentro_q=False)
            return
    # tabla con solo imágenes (ideas de diagramas): opciones de imagen
    if not boton and ncol == 1 and all(any(b[0] == "p" and b[1]["imgs"] for b in c) for f in filas for c in f):
        q = get_q()
        if q is not None:
            for f in filas:
                for c in f:
                    for b in c:
                        if b[0] == "p":
                            for u in b[1]["imgs"]:
                                q.setdefault("opciones_img", []).append(u)
            return
    poner(dict(k="tabla", filas=filas))


# ------------------------------------------------------------------ render


def render_img(u, cls="fig"):
    return '<img class="%s" src="%s" alt="">' % (cls, u)


def render_bloques_celda(bloques, ctx=None):
    h = []
    pend_ul = None
    previo_marca = False
    modo_opts = False
    for tipo, b in bloques:
        if tipo == "t":
            h.append(render_tabla(b))
            continue
        for u in b["imgs"]:
            h.append(render_img(u))
        lins = lineas_de(b, True)
        if b["num"] and b["num"]["fmt"] == "bullet":
            texto = " ".join(x[0] for x in lins)
            h.append("__LI__" + (("chk:" if previo_marca else "") + texto))
            continue
        for lt, lp in lins:
            if modo_opts and b["neg"] < 0.5:
                h.append("__OP__" + lt)
                continue
            modo_opts = bool(re.search(r"[Mm]arca (todas|las alternativas|cada)", lp)) and "(V)" not in lp
            previo_marca = bool(re.search(r"[Mm]arca", lp))
            h.append("<p>%s</p>" % inline(lt))
    # agrupar items de lista
    out, i = [], 0
    while i < len(h):
        if h[i].startswith("__OP__"):
            items = []
            while i < len(h) and h[i].startswith("__OP__"):
                items.append(h[i][6:])
                i += 1
            out.append("<ul class='opts'>%s</ul>" % "".join(
                "<li><label><input type='checkbox'><span>%s</span></label></li>" % inline(x) for x in items))
        elif h[i].startswith("__LI__"):
            items = []
            chk = h[i].startswith("__LI__chk:")
            while i < len(h) and h[i].startswith("__LI__"):
                items.append(h[i][6:].replace("chk:", "", 1) if h[i][6:].startswith("chk:") else h[i][6:])
                i += 1
            if chk:
                out.append("<ul class='opts chk'>%s</ul>" % "".join("<li><span>%s</span></li>" % inline(x) for x in items))
            else:
                out.append("<ul class='l'>%s</ul>" % "".join("<li>%s</li>" % inline(x) for x in items))
        else:
            out.append(h[i])
            i += 1
    return "".join(out)


_uid = [0]


def render_tabla(filas):
    _uid[0] += 1
    g = "g%d" % _uid[0]
    if es_vf(filas):
        h = ["<table class='vf'><tr><th class='izq'>Afirmación</th><th>V</th><th>F</th></tr>"]
        for r, f in enumerate(filas[1:]):
            enun = render_bloques_celda(f[0]) if f and f[0] else ""
            h.append(fila_vf(enun, "%s_%d" % (g, r)))
        h.append("</table>")
        return "".join(h)
    # figura (celdas con imágenes)
    if any(b[0] == "p" and b[1]["imgs"] for f in filas for c in f for b in c):
        h = ["<div class='figs'>"]
        for f in filas:
            for c in f:
                if c:
                    h.append("<figure>%s</figure>" % render_bloques_celda(c))
        h.append("</div>")
        return "".join(h)
    h = ["<table>"]
    for r, f in enumerate(filas):
        h.append("<tr>")
        for c in f:
            t = texto_celda(c)
            m = re.fullmatch(r"\{(.+)\}", t.strip())
            tag = "th" if r == 0 and not m and not re.fullmatch(r"\W*Completa\W*", t) and len(filas) > 1 and all(
                texto_celda(x) for x in filas[0]) and not any(re.fullmatch(r"\{.*\}", texto_celda(y)) for y in filas[0]) else "td"
            if m:
                marcado = " ".join(b[1]["texto"].strip() for b in c if b[0] == "p").strip()
                mm = re.fullmatch(r"\{(.+)\}", marcado)
                cell = control_select(opciones_de(mm.group(1) if mm else m.group(1)))
            elif re.fullmatch(r"\W*Completa\W*", t):
                cell = control_input()
            else:
                cell = render_bloques_celda(c)
            h.append("<%s>%s</%s>" % (tag, cell, tag))
        h.append("</tr>")
    h.append("</table>")
    return "".join(h)


def fila_vf(enun_html, nombre):
    _, atr = nuevo_control("vf")
    return ("<tr%s><td class='izq'>%s</td><td><input type='radio' name='%s' value='V' aria-label='Verdadero'></td>"
            "<td><input type='radio' name='%s' value='F' aria-label='Falso'></td></tr>") % (atr, enun_html, nombre, nombre)


def render_diapositiva(celda_bloques):
    """Una celda -> (titulo, html)."""
    bl = list(celda_bloques)
    # quitar 'Diapositiva N'
    while bl and bl[0][0] == "p" and re.match(r"^\W*Diapositiva\s*\d+\W*$", bl[0][1]["plano"].strip()):
        bl.pop(0)
    titulo = ""
    if bl and bl[0][0] == "p" and bl[0][1]["neg"] >= 0.8 and len(bl[0][1]["plano"]) < 90:
        titulo = bl[0][1]["plano"].strip()
        bl.pop(0)
    return titulo, render_bloques_celda(bl)


def render_carrusel(it, cid):
    slides = it.get("slides") or [render_diapositiva(c) for c in it.get("filas", [])]
    h = ["<div class='car' id='%s' data-n='%d'><div class='car-et'>%s</div>" % (cid, len(slides), html.escape(it["etiqueta"]))]
    for i, (tit, cuerpo) in enumerate(slides):
        h.append("<div class='sl%s'><div class='sl-n'>%d / %d</div>%s<div class='sl-b'>%s</div></div>" % (
            " on" if i == 0 else "", i + 1, len(slides), ("<h3>%s</h3>" % inline(tit)) if tit else "", cuerpo))
    if len(slides) == 1:
        h.append("</div>")
        return "".join(h).replace("<div class='car'", "<div class='car uno'", 1)
    h.append("<div class='car-nav'><button type='button' class='prev' aria-label='Anterior'>‹</button>"
             "<span class='dots'>%s</span><button type='button' class='next' aria-label='Siguiente'>›</button></div></div>" % "".join(
                 "<i class='%s'></i>" % ("on" if i == 0 else "") for i in range(len(slides))))
    return "".join(h)


def render_explore(titulo, cuerpo):
    """Exploremos plegado y bloqueado: se abre al comprobar la pregunta anterior."""
    return ("<details class='exp bloq'><summary><span class='et'>%s</span><span class='cand'>Se muestra después de responder</span>"
            "</summary><div>%s</div></details>") % (titulo or "Exploremos una posible respuesta", cuerpo)


def render_item(it, cid, personajes, orden):
    if it.get("tras"):
        h = render_item(dict(it, tras=False), cid, personajes, orden)
        return "<div class='tras'>%s</div>" % h if h else ""
    k = it["k"]
    if k == "p":
        return "<p class='txt'>%s</p>" % inline(it["t"])
    if k == "titulo":
        return "<h3 class='sub'>%s</h3>" % inline(it["t"])
    if k == "ul":
        return "<ul class='l'>%s</ul>" % "".join("<li>%s</li>" % inline(x) for x in it["items"])
    if k == "img":
        return render_img(it["uri"])
    if k == "objetivo":
        return "<div class='obj'><strong>%s</strong>%s</div>" % (inline(it["lineas"][0]), "".join("<p>%s</p>" % inline(x) for x in it["lineas"][1:]))
    if k == "dialogo":
        h = ["<div class='dia'>"]
        for nom, txt in it["lineas"]:
            if nom not in orden:
                orden.append(nom)
            lado = "izq" if orden.index(nom) % 2 == 0 else "der"
            h.append("<div class='fila %s'><div class='avc'>%s<span class='nom'>%s</span></div><div class='burb'>%s</div></div>" % (
                lado, avatar_svg(nom, orden.index(nom)), html.escape(nom), inline(txt)))
        h.append("</div>")
        return "".join(h)
    if k == "carrusel":
        return render_carrusel(it, cid)
    if k == "explore":
        tit, cuerpo = render_diapositiva(it["celda"])
        return render_explore(inline(tit) if tit else "", cuerpo)
    if k == "exph":
        return render_explore(it["titulo"], it["html"])
    if k == "html":
        return it["html"]
    if k == "qmd":
        return envolver_pregunta(it["n"], it["enun"], it["html"], it["claves"])
    if k == "caja":
        tit, cuerpo = render_diapositiva(it["celda"])
        return "<div class='caja'>%s%s</div>" % (("<h4>%s</h4>" % inline(tit)) if tit else "", cuerpo)
    if k == "tabla":
        return render_tabla(it["filas"])
    if k == "boton":
        cont = "".join(render_item(c, cid + "b", personajes, orden) for c in it["contenido"])
        return "<details class='btn'><summary>%s</summary>%s</details>" % (html.escape(it["texto"]), cont)
    if k == "recurso":
        return "<div class='rec'><span>Recurso interactivo por producir</span>%s</div>" % html.escape(it["t"])
    if k == "q":
        multi = it["multi"] or len(re.findall(r"(?i)\b[a-e]\b", " ".join(it.get("claves") or [])[:20])) > 1 and bool(it["opciones"])
        _uid[0] += 1
        g = "q%d" % _uid[0]
        enun = "<p class='enun'>%s</p>" % inline(it["stem"])
        h = []
        for b in it["bloques"]:
            h.append(render_item(b, cid, personajes, orden))
        if it.get("opciones_img"):
            h.append(ul_alternativas(it["opciones_img"], False, g, imagenes=True))
        if it["opciones"] and any(re.search(r"_{3,}", o) for o in it["opciones"]):
            # sublista numerada con campos: son partes de la pregunta, no alternativas
            h.extend("<p class='parte'><span class='pl'>%s.</span> %s</p>" % ("abcdefgh"[j % 8], inline(o))
                     for j, o in enumerate(it["opciones"]))
        elif it["opciones"]:
            h.append(ul_alternativas(it["opciones"], multi, g))
        return envolver_pregunta(it["n"], enun, "".join(h), it.get("claves"))
    return ""


# ------------------------------------------------------------------ lectura del .md
#
# El .md se lee directo, sin pasar por guion_a_docx.py: ese .docx es el de producción (cajas de una
# celda, colores por marca) y no conserva las listas ni separa diapositivas, preguntas y alternativas.
# El lector arma los mismos ítems que parsear() para el .docx (Comp con items: p, dialogo, carrusel,
# exph, tabla, img…) y cada pregunta como un ítem "qmd" con su HTML y sus claves.

RE_PREG = (re.compile(r"^\*\*(\d{1,2})\.\*\*\s+(.+)$"),           # **4.** enunciado
           re.compile(r"^(\d{1,2})\.\s+\*\*(.+?)\*\*\s*$"),        # 1. **enunciado**
           re.compile(r"^\*\*(\d{1,2})\.\s+(.+?)\*\*\s*$"))        # **28. enunciado**
RE_NUM = re.compile(r"^(\*\*)?(\d)[.)](\*\*)?\s+(.*)$")             # 1. / 1) alternativa o parte
RE_PARTE = re.compile(r"^(\*{0,3})([a-d])(\*{0,2})\s*([.)])(\*{0,2})\s*\.?(?=\s|\*|$)\s*(.*)$")
RE_STATIC = re.compile(r"https?://static\.rutasdigitales[^\s)\"'<>]+")
RE_DRIVE = re.compile(r"https?://drive\.google\.com/[^\s)\"'<>]+")
RE_CORTE = re.compile(r"(?i)^\W*propuesta de preguntas para evaluaci|^\W*preguntas (para la )?evaluaci")


def preparar_md(texto):
    """Normaliza el .md: citas, escapes de doc_a_md, $`…`$, fórmulas en varias líneas."""
    out = []
    for l in texto.split("\n"):
        l = l.rstrip()
        l = re.sub(r"^(\s*>)+\s?", "", l)
        if l.endswith("\\") and not l.endswith("\\\\"):
            l = l[:-1].rstrip()
        l = l.replace("<!-- -->", "")
        l = re.sub(r"\$`(.+?)`\$", r"\\(\1\\)", l)
        l = re.sub(r"^(\s*\**\s*)([a-dA-D1-9])(\**)\\\)", r"\1\2\3)", l)       # a\) -> a)
        for a, b in (("\\\\[", "\\["), ("\\\\]", "\\]"), ("\\\\(", "\\("), ("\\\\)", "\\)")):
            l = l.replace(a, b)
        # fórmula escrita sin delimitadores ("- V(10,3)=\text{___}\cdot…"): se envuelve en \( \)
        m = re.match(r"^(\s*-\s+|\s*)(.*)$", l)
        if m and not re.search(r"\\[(\[]|\$|^\s*IP", m.group(2)) and \
                re.search(r"\\(cdot|text|frac|dfrac|times|div|binom)\b", m.group(2)) and \
                not re.search(r"[a-záéíóúñ]{4,}\s+[a-záéíóúñ]{3,}\s", re.sub(r"\\\w+\{[^{}]*\}|\\\w+", "", m.group(2))):
            l = m.group(1) + "\\(" + m.group(2).strip() + "\\)"
        if re.match(r"^(Mostrar|Después de validar|Luego de validar)\b", l):
            l = "IPP: " + l                                  # instrucciones sin marca
        out.append(l)
    # fórmula de bloque partida en varias líneas: \[ … \]
    res, i = [], 0
    while i < len(out):
        l = out[i]
        if re.search(r"\\\[(?!.*\\\])", l) and not RE_SEL.search(l):
            j, junto = i + 1, l
            while j < len(out) and j < i + 12 and "\\]" not in junto[junto.rfind("\\["):]:
                if out[j].strip():
                    junto += " " + out[j].strip()
                j += 1
            if "\\]" in junto[junto.rfind("\\["):]:
                res.append(junto)
                res.extend([""] * (j - i - 1))
                i = j
                continue
        res.append(l)
        i += 1
    return res


def plano_md(l):
    return re.sub(r"^#+\s*", "", l.strip()).replace("*", "").strip()


def es_marca(l, *prefijos):
    p = plano_md(l).lower()
    return any(p.startswith(x) for x in prefijos)


def es_negrita(l):
    s = l.strip()
    return s.startswith("**") and s.rstrip(". ").endswith("**") and len(s) > 4


def es_dialogo_ini(l):
    return re.match(r"(?i)^ipp:\s*inicio (de )?di[aá]logo", plano_md(l)) is not None


def es_dialogo_fin(l):
    return re.match(r"(?i)^ipp:\s*fin (de )?di[aá]logo", plano_md(l)) is not None


def arbol_html(js, nombre):
    """Árbol del plugin (JSON) dibujado con render_arbol.py, como imagen incrustada."""
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        import render_arbol as RA
        fig = RA.dibujar(json.loads(js), nombre, None, 1.0, titulo=False)
        buf = io.BytesIO()
        fig.savefig(buf, format="png", facecolor=fig.get_facecolor())
        plt.close(fig)
        b = base64.b64encode(buf.getvalue()).decode()
        return "<div class='tree'><img class='fig' alt='%s' src='data:image/png;base64,%s'></div>" % (
            html.escape(nombre or "Diagrama de árbol"), b)
    except Exception as e:  # un JSON roto no debe impedir la vista
        return "<div class='ph'><span>Diagrama de árbol %s</span>no se pudo dibujar: %s</div>" % (
            html.escape(nombre), html.escape(str(e)))


def es_capsula(l):
    p = plano_md(l).rstrip(".").lower()
    return p in ("cápsula de contenido", "ipp: cápsula de contenido", "cápsula", "ipp: cápsula",
                 "capsula de contenido", "ipp: capsula de contenido",
                 "cápsula de observación", "ipp: cápsula de observación",
                 "capsula de observacion", "ipp: capsula de observacion")


def es_exploremos(l):
    return re.match(r"(?i)^(exploremos|analicemos)\b", plano_md(l)) is not None


def es_diapositiva(l):
    return re.match(r"^Diapositiva\s*\d+", plano_md(l)) is not None


def es_tabla(l):
    return l.strip().startswith("|")


def es_hueco(l):
    """¿La línea pide una respuesta? (campo ___, caja \\boxed{\\qquad}, "Escribe tu respuesta").
    Las notas IPx no cuentan, aunque citen un campo."""
    s = l.strip()
    if re.fullmatch(r"[*_\s]+", s) or re.match(r"^\**\s*IP[PDMC]\b", s):
        return False
    return bool(re.search(r"_{3,}|\\boxed\{\s*\\q?quad|^\W*Escribe tu respuesta", s))


def pregunta_en(l):
    for r in RE_PREG:
        m = r.match(l.strip())
        if m:
            return int(m.group(1)), m.group(2).strip()
    return None


def sin_negrita(t):
    t = t.strip()
    if t.startswith("**") and t.endswith("**") and t.count("**") == 2:
        return t[2:-2].strip()
    if t.count("**") % 2:
        t = t[:-2] if t.endswith("**") else t.replace("**", "", 1)
    return t


def leer_tabla_md(lineas, i):
    filas, cab_vacia = [], None
    while i < len(lineas) and es_tabla(lineas[i]):
        l = lineas[i].strip()
        if not re.fullmatch(r"\|[\s:|-]+\|", l):
            celdas = [c.strip() for c in l.strip("|").split("|")]
            if cab_vacia is None:
                cab_vacia = not any(c for c in celdas)
            filas.append(celdas)
        i += 1
    return filas, bool(cab_vacia), i


class LectorMd:
    def __init__(self, ruta):
        self.ruta = ruta
        self.L = preparar_md(encabezado_corto.normalizar(open(ruta, encoding="utf-8").read()))
        self.alts = alternativas.detectar(self.L, avisar=False)
        self.pregs = self.buscar_preguntas()
        self.refs = {}            # TAB_… / IMG_… -> html, para los botones "Ver …"
        self.unalinea = False     # "IPP: Tabla TAB_… … una sola línea": la próxima tabla no parte sus casillas
        self.meta = {}
        self.uid = 0

    def buscar_preguntas(self):
        """{índice de línea: (número, enunciado)}. Además de las formas en negrita, acepta "20. texto"
        sin negrita cuando sigue la numeración (20 después de 19)."""
        out, ultimo = {}, 0
        for k, l in enumerate(self.L):
            r = pregunta_en(l)
            if r is None:
                m = re.match(r"^(\d{1,2})\.\s+(\S.*)$", l.strip())
                if m and int(m.group(1)) == ultimo + 1 and int(m.group(1)) >= 5:
                    r = int(m.group(1)), m.group(2).strip()
            if r:
                out[k] = r
                ultimo = r[0]
        return out

    # ---------------------------------------------------------------- piezas
    def img(self, url, alt="", nombre=None, tras=False):
        return dict(k="html", nombre=nombre, tras=tras,
                    html="<img class='fig' src='%s' alt='%s' loading='lazy'>" % (html.escape(url), html.escape(alt)))

    def ph(self, texto, nombre=None, tras=False):
        return dict(k="html", nombre=nombre, tras=tras,
                    html="<div class='ph'><span>Imagen pendiente</span>%s</div>" % html.escape(texto))

    def tabla(self, filas, cab_vacia, multi_filas=False):
        while filas and not any(filas[0]):
            filas = filas[1:]
        if not filas:
            return ""
        cab = [plano_md(c) for c in filas[0]]
        # verdadero/falso: columna "Respuesta" con "___"
        if len(filas) > 1 and cab and cab[-1].lower() in ("respuesta", "respuestas", "v o f", "v/f") and all(
                re.fullmatch(r"[*_\s]*", f[-1]) for f in filas[1:]):
            self.uid += 1
            h = ["<table class='vf'><tr><th class='izq'>%s</th><th>V</th><th>F</th></tr>" % inline(filas[0][0], False)]
            for r, f in enumerate(filas[1:]):
                h.append(fila_vf(inline(f[0]), "vf%d_%d" % (self.uid, r)))
            return "".join(h) + "</table>"
        h = []
        if multi_filas:
            _, atr = nuevo_control("multi", filas=True)
            h.append("<table class='selfilas'%s>" % atr)
        else:
            h.append("<table>")
        for r, f in enumerate(filas):
            th = r == 0 and len(filas) > 1 and (cab_vacia and all(es_negrita(c) or not c for c in f) or not cab_vacia)
            tag = "th" if th else "td"
            h.append("<tr>")
            if multi_filas:
                h.append("<th></th>" if th else "<td><input type='checkbox' value='%s' aria-label='Marcar'></td>"
                         % html.escape(plano_md(f[0])))
            for c in f:
                m = re.fullmatch(r"\{(.+)\}", c.strip())
                if m and not th:
                    cel = control_select(opciones_de(m.group(1)))
                elif re.fullmatch(r"\W*Completa\W*", c) and not th:
                    cel = control_input()
                else:
                    cel = inline(c if th else c, huecos=not th)
                    if th:
                        cel = cel.replace("<strong>", "").replace("</strong>", "")
                h.append("<%s>%s</%s>" % (tag, cel, tag))
            h.append("</tr>")
        return "".join(h) + "</table>"

    def dialogo(self, lineas):
        out = []
        for l in lineas:
            s = l.strip()
            if not s or plano_md(s) in ("", "*"):
                continue
            m = re.match(r"^\**\s*([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+(?: [A-ZÁÉÍÓÚÑ][a-záéíóúñ]+)?)\s*\**\s*:\s*\**\s*(.*)$", s)
            if m and m.group(1) not in NO_PERSONAJES:
                out.append([m.group(1), sin_negrita(m.group(2))])
            elif out:
                out[-1][1] += " " + sin_negrita(s)
        return dict(k="dialogo", lineas=[tuple(x) for x in out]) if out else None

    def nota_ipp(self, l, items):
        """Nota de producción: casi siempre se oculta; de algunas sale una imagen, un recurso o un botón."""
        p = plano_md(l)
        low = p.lower()
        tras = bool(re.search(r"(?i)(despu[eé]s|luego) de (validar|comprobar|enviar)", p))
        m = re.match(r"(?i)^ipp:\s*tabla de referencia\s+(\w+)", p) or re.match(r"(?i)^ipp:\s*tabla\s+(TAB_\w+)", p)
        if m:
            self.unalinea = bool(re.search(r"(?i)una sola l[ií]nea", p))
            return ("ref", m.group(1))
        m = re.match(r"(?i)^ipp:\s*bot[oó]n\s*[“\"](.+?)[”\"]\s*[,.]?\s*mostrar\s+(?:el\s+diagrama\s+)?(\w+)", p)
        if m:
            cont = self.refs.get(m.group(2)) or "<p class='txt'>(%s)</p>" % html.escape(m.group(2))
            items.append(dict(k="html", html="<details class='btn'><summary>%s</summary>%s</details>" % (
                html.escape(m.group(1)), cont)))
            return None
        u = RE_STATIC.search(l)
        if u and re.search(r"\.html?\b", u.group(0)):
            url = u.group(0)
            items.append(dict(k="html", tras=tras, html=(
                "<div class='recw'><iframe src='%s' title='Recurso interactivo'></iframe>"
                "<a href='%s' target='_blank' rel='noopener'>Abrir el recurso en otra pestaña</a></div>") % (
                html.escape(url), html.escape(url))))
            return None
        if u and re.search(r"\.(png|jpe?g|gif|svg|webp)$", u.group(0), re.I):
            nom = re.search(r"\b(IMG_\w+)", p)
            alt = re.sub(r"\(\[link\]\(.*?\)\)|\[link\]\(.*?\)", "", re.sub(r"^IP[PD]:\s*", "", p)).strip()
            it = self.img(u.group(0), alt, nom.group(1) if nom else None, tras)
            items.append(it)
            if nom:
                self.refs[nom.group(1)] = it["html"]
            return None
        if RE_DRIVE.search(l) and re.search(r"(?i)imagen", p):
            nom = re.search(r"\b(IMG_\w+)", p)
            texto = "Imagen en Google Drive, sin subir al servidor%s." % ((" (%s)" % nom.group(1)) if nom else "")
            if nom:
                self.refs[nom.group(1)] = "<div class='ph'><span>Imagen pendiente</span>%s</div>" % html.escape(texto)
            if tras or re.search(r"(?i)mostrar imagen", p):
                items.append(self.ph(texto, tras=tras))
            return None
        if tras:
            m = re.search(r"\b(IMG_\w+)", p)
            if m:
                for it in items:
                    if it.get("nombre") == m.group(1):
                        it["tras"] = True
        if low.startswith("ipc:") and re.search(r"recurso|diagrama|imagen", low) and "ideas" not in low:
            items.append(dict(k="recurso", t=re.sub(r"(?i)^IPC:\s*", "", p)))
        return None

    # ---------------------------------------------------------------- bloque de líneas
    def bloques(self, i0, i1, en_pregunta=False, claves=None):
        """Lee las líneas [i0, i1) y devuelve una lista de ítems."""
        L = self.L
        items = []
        i = i0
        enun_prev = ""
        ref_pend = None
        clave_multilinea = False
        nota_dos_puntos = False
        while i < i1:
            l = L[i]
            s = l.strip()
            p = plano_md(l)
            if not p or re.fullmatch(r"[_*.\s-]*", s) and s != "---":
                i += 1
                continue
            if s == "---":
                i += 1
                continue
            # claves (no se muestran)
            if en_pregunta:
                if clave_multilinea and RE_PARTE.match(s.replace("**", "")):
                    claves.extend(tokens_clave("Respuestas: " + s) or [])
                    i += 1
                    continue
                clave_multilinea = False
                toks = tokens_clave(s) if (p.startswith("IPP") or re.match(r"(?i)^(respuestas?|correctas?)\b", p)) else None
                if toks is not None:
                    claves.extend(toks)
                    clave_multilinea = not toks
                    i += 1
                    continue
            if nota_dos_puntos and re.fullmatch(r"\\\[.*\\\]\.?", s):
                i += 1                                     # sigue siendo parte de la nota: no se muestra
                continue
            nota_dos_puntos = False
            if es_dialogo_ini(l):
                j = i + 1
                while j < i1 and not es_dialogo_fin(L[j]):
                    j += 1
                d = self.dialogo(L[i + 1:j])
                if d:
                    items.append(d)
                i = j + 1
                continue
            if es_dialogo_fin(l):
                i += 1
                continue
            if es_tabla(l):
                filas, cab_vacia, j = leer_tabla_md(L, i)
                multi = en_pregunta and any(re.search(r"(?i)selecci[oó]n m[uú]ltiple", L[k]) for k in range(j, i1)) and \
                    all(re.fullmatch(r"\**\d+\**", c[0].strip()) for c in filas[1:] if c) and \
                    not any("data-tipo='multi'" in x.get("html", "") for x in items)
                h = self.tabla(filas, cab_vacia, multi_filas=multi)
                if self.unalinea:
                    h = "<div class='tabw'>%s</div>" % h.replace("<table", "<table data-unalinea='1'", 1)
                    self.unalinea = False
                if not (ref_pend and ref_pend.upper().endswith("_ANS")):   # TAB_…_ANS: solo en los botones "Ver …"
                    items.append(dict(k="html", html=h))
                if ref_pend:
                    self.refs[ref_pend] = h
                    ref_pend = None
                i = j
                continue
            if es_capsula(l) or es_diapositiva(l):
                j, it = self.capsula(i, i1)
                items.append(it)
                i = j
                continue
            u_sola = RE_STATIC.fullmatch(p.strip())
            if u_sola and re.search(r"\.html?\b", u_sola.group(0)):
                # URL de un recurso en su propia línea, después de "IPP: incrustar recurso html…"
                self.nota_ipp(u_sola.group(0), items)
                i += 1
                continue
            if p.startswith(("IPP", "IPD", "IPM", "IPC")) or p == "IPP":
                nota_dos_puntos = p.endswith(":")
                r = self.nota_ipp(l, items)
                if r and r[0] == "ref":
                    ref_pend = r[1]
                i += 1
                continue
            m = re.match(r"^\[imagen:\s*(.+?)\]$", p)
            if m:
                items.append(self.ph("Imagen del Doc (%s); no está en el .md." % m.group(1)))
                i += 1
                continue
            if es_capsula(l) or es_diapositiva(l):
                j, it = self.capsula(i, i1)
                items.append(it)
                i = j
                continue
            if es_exploremos(l):
                j, it = self.exploremos(i, i1)
                items.append(it)
                i = j
                continue
            if re.match(r"(?i)^recordemos", p) and len(p) < 20:
                items.append(dict(k="titulo", t="Recordemos…"))
                self.recordemos = True
                i += 1
                continue
            if s.startswith("#"):
                items.append(dict(k="titulo", t=p))
                i += 1
                continue
            # árboles del plugin: la línea JSON que sigue a "Referencia DIAG_Pnn". El _ANS no se muestra.
            if s.startswith('{"mode"'):
                k = i - 1
                while k >= 0 and (not L[k].strip() or L[k].strip().startswith("```")):
                    k -= 1
                previa = L[k] if k >= 0 else ""
                mref = re.search(r"DIAG_P\w+", previa)
                nombre = mref.group(0) if mref else ""
                if not (nombre.endswith("_ANS") or "respuesta" in previa.lower()):
                    items.append(dict(k="html", html=arbol_html(s, nombre)))
                elif nombre:                       # el _ANS solo en los botones "Ver árbol"
                    self.refs[nombre] = arbol_html(s, nombre)
                i += 1
                continue
            if s.startswith("```"):
                i += 1
                continue
            if re.match(r"(?i)^escribe tu respuesta", p):
                items.append(dict(k="html", html="<textarea class='abierta' rows='4' placeholder='Escribe tu respuesta'></textarea>"))
                i += 1
                continue
            # alternativas sin numerar (alternativas.py)
            if i in self.alts:
                textos = []
                while i < i1 and (i in self.alts or not L[i].strip()):
                    if L[i].strip():
                        textos.append(alternativas.quitar_prefijo(L[i].strip()))
                    i += 1
                self.uid += 1
                items.append(dict(k="html", html=ul_alternativas(textos, self.es_multi(enun_prev), "a%d" % self.uid)))
                continue
            # partes a. / **b)**
            mp = RE_PARTE.match(s)
            if mp and en_pregunta:
                letra, resto = mp.group(2), mp.group(6)
                neg = bool(mp.group(1)) or resto.rstrip(". ").endswith("**")
                resto = sin_negrita(resto.replace("**", "") if neg else resto)
                enun_prev = resto
                items.append(dict(k="html", html="<p class='parte%s'>\x05%s\x05<span class='pl'>%s.</span> %s</p>" % (
                    " enun" if neg else "", letra, letra, inline(resto))))
                i += 1
                continue
            # ítems numerados: alternativas o partes
            mn = RE_NUM.match(s)
            if mn and en_pregunta:
                j, it = self.numerados(i, i1, enun_prev)
                items.extend(it)
                if it and it[0].get("enun"):
                    enun_prev = it[0]["enun"]
                i = j
                continue
            if s.startswith("- "):
                lis = []
                while i < i1 and L[i].strip().startswith("- "):
                    lis.append(L[i].strip()[2:])
                    i += 1
                items.append(dict(k="html", html="<ul class='l'>%s</ul>" % "".join("<li>%s</li>" % inline(x) for x in lis)))
                continue
            if en_pregunta and es_negrita(s):
                enun_prev = p
                items.append(dict(k="html", html="<p class='enun'>%s</p>" % inline(sin_negrita(s))))
            else:
                if p.endswith(("?", ":")):
                    enun_prev = p
                items.append(dict(k="p", t=s))
            i += 1
        return items

    def es_multi(self, enun):
        return bool(re.search(r"(?i)(marca|selecciona) (todas|las alternativas|cada)\b(?!.*verdader)", enun))

    def numerados(self, i, i1, enun_prev):
        """Ítem "1." / "1)": alternativa si va seguido de otra o de la clave; parte si trae contenido propio."""
        L = self.L

        def siguiente(jj):
            sig = jj + 1
            while sig < i1 and not L[sig].strip():
                sig += 1
            return sig

        def es_alt(jj):
            m = RE_NUM.match(L[jj].strip())
            if not m or m.group(1) or es_hueco(m.group(4)) or es_negrita(m.group(4)):
                return False
            sig = siguiente(jj)
            ps = L[sig].strip() if sig < i1 else ""
            return bool(sig >= i1 or RE_NUM.match(ps) or plano_md(ps).startswith("IP") or RE_PARTE.match(ps)
                        or es_negrita(ps) or re.fullmatch(r"[_*.\s]*", ps) or es_exploremos(ps)
                        or es_dialogo_ini(ps) or re.match(r"(?i)^(respuestas?|correctas?)\b", plano_md(ps)))

        if es_alt(i):
            grupo, j = [], i
            while j < i1 and RE_NUM.match(L[j].strip()) and es_alt(j):
                grupo.append(RE_NUM.match(L[j].strip()).group(4))
                j = siguiente(j)
            if len(grupo) >= 2:
                self.uid += 1
                return j, [dict(k="html", html=ul_alternativas(grupo, self.es_multi(enun_prev), "n%d" % self.uid))]
            return j, [dict(k="p", t=grupo[0])]
        m = RE_NUM.match(L[i].strip())
        letra = "abcdefgh"[int(m.group(2)) - 1]
        texto = m.group(4)
        neg = bool(m.group(1)) or es_negrita(texto)
        return i + 1, [dict(k="html", enun=plano_md(texto), html="<p class='parte%s'>\x05%s\x05<span class='pl'>%s.</span> %s</p>" % (
            " enun" if neg else "", letra, letra, inline(sin_negrita(texto))))]

    def capsula(self, i, i1):
        """Cápsula o serie de diapositivas -> carrusel. Termina en '---', una nota IPP o una pregunta."""
        L = self.L
        etiqueta = "Diapositivas"
        if es_capsula(L[i]):
            if "observ" in plano_md(L[i]).lower():
                etiqueta = "Cápsula de observación"
            else:
                etiqueta = "Recordemos" if getattr(self, "recordemos", False) else "Cápsula de contenido"
            i += 1
        self.recordemos = False
        slides, tit, cuerpo = [], "", []

        def cerrar():
            if tit or any(x.strip() for x in cuerpo):
                slides.append((tit, cuerpo[:]))

        j = i
        while j < i1:
            l = L[j]
            s = l.strip()
            p = plano_md(l)
            # Dentro de una diapositiva: el código de una tabla que viene justo debajo (IPP: Tabla TAB_…) y la IPD de
            # una imagen que viene justo debajo son parte de ella.
            sig = next((L[k].strip() for k in range(j + 1, i1) if L[k].strip()), "")
            dentro = (re.match(r"(?i)^ipp:\s*tabla\s+TAB_", p) and sig.startswith("|")) or \
                (re.match(r"(?i)^ipd:", p) and re.match(r"(?i)^ipp:.*imagen", plano_md(sig)) and RE_STATIC.search(sig)) or \
                (re.match(r"(?i)^ipp:\s*programar como serie de diapositivas", p) and not (tit or cuerpo))
            if s == "---" or es_dialogo_ini(l) or es_exploremos(l) or j in self.pregs or \
                    (p.startswith(("IPP", "IPD", "IPM", "IPC")) and not RE_STATIC.search(l) and not dentro) or \
                    (es_capsula(l) and (tit or cuerpo)) or re.fullmatch(r"_{5,}", s):
                break
            if dentro and re.match(r"(?i)^ipp:\s*programar", p):
                j += 1
                continue
            if es_diapositiva(l):
                cerrar()
                tit, cuerpo = re.sub(r"^Diapositiva\s*\d+\s*[:.\-–]?\s*", "", p).strip(), []
                j += 1
                continue
            if not p or p == "*":
                j += 1
                continue
            if es_negrita(s) and len(p) < 100 and not p.startswith("\\"):
                if not tit and not any(x.strip() for x in cuerpo):
                    tit = p
                    j += 1
                    continue
                if any(x.strip() for x in cuerpo):
                    cerrar()
                    tit, cuerpo = p, []
                    j += 1
                    continue
            cuerpo.append(l)
            j += 1
        cerrar()
        out = []
        for t, c in slides:
            out.append((t, self.html_simple(c)))
        return j, dict(k="carrusel", etiqueta=etiqueta, slides=out)

    def html_simple(self, lineas):
        """Cuerpo de diapositiva o de Exploremos: párrafos, tablas, listas e imágenes."""
        h, k = [], 0
        while k < len(lineas):
            l = lineas[k]
            s = l.strip()
            p = plano_md(l)
            if not p or re.fullmatch(r"[_*\s]+", s):
                k += 1
                continue
            if es_tabla(l):
                filas, cab_vacia, k2 = leer_tabla_md(lineas, k)
                h.append(self.tabla(filas, cab_vacia))
                k = k2
                continue
            u = RE_STATIC.search(l)
            if p.startswith(("IPP", "IPD")):
                if u and re.search(r"\.(png|jpe?g|gif|svg|webp)$", u.group(0), re.I):
                    h.append("<img class='fig' src='%s' alt='' loading='lazy'>" % html.escape(u.group(0)))
                k += 1
                continue
            if s.startswith("- ") or re.match(r"^\d+[.)]\s", s):
                ordenada = not s.startswith("- ")
                lis = []
                while k < len(lineas) and (lineas[k].strip().startswith("- ") if not ordenada else
                                           re.match(r"^\d+[.)]\s", lineas[k].strip()) or not lineas[k].strip()):
                    if lineas[k].strip():
                        lis.append(re.sub(r"^(- |\d+[.)]\s+)", "", lineas[k].strip()))
                    k += 1
                tag = "ol" if ordenada else "ul"
                h.append("<%s class='l'>%s</%s>" % (tag, "".join("<li>%s</li>" % inline(x) for x in lis), tag))
                continue
            h.append("<p>%s</p>" % inline(s))
            k += 1
        return "".join(h)

    def exploremos(self, i, i1):
        L = self.L
        titulo = plano_md(L[i])
        j = i + 1
        while j < i1 and not L[j].strip():
            j += 1
        cuerpo = []
        if j < i1 and re.fullmatch(r"\**_{5,}\**", L[j].strip()):       # recuadro de rayas
            j += 1
            while j < i1 and not re.fullmatch(r"\**_{5,}\**", L[j].strip()):
                cuerpo.append(L[j])
                j += 1
            j += 1
        else:
            while j < i1:
                l = L[j]
                p = plano_md(l)
                if j in self.pregs or es_dialogo_ini(l) or es_capsula(l) or l.strip() == "---" or \
                        (p.startswith(("IPP", "IPM", "IPC")) and not RE_STATIC.search(l)) or \
                        (p.startswith("IPD") and not RE_STATIC.search(l)) or RE_PARTE.match(l.strip()):
                    break
                cuerpo.append(l)
                j += 1
        return j, dict(k="exph", titulo=html.escape(titulo), html=self.html_simple(cuerpo))

    # ---------------------------------------------------------------- lección completa
    def leer(self):
        L = self.L
        n = len(L)
        fin = next((k for k in range(n) if RE_CORTE.match(plano_md(L[k])) and es_negrita(L[k])), n)
        comps, cur = [], None
        meta = {}
        i = 0
        # encabezado: tablas antes del primer IPM
        while i < fin and not es_marca(L[i], "ipm:"):
            if es_tabla(L[i]):
                filas, _, i = leer_tabla_md(L, i)
                for f in filas:
                    if len(f) >= 2:
                        k, v = plano_md(f[0]), plano_md(f[1])
                        if k.startswith("Título"):
                            meta["titulo"] = v
                        elif k.startswith("Objetivo"):
                            meta["obj"] = v
                        elif k in ("Unidad", "Tema", "Curso"):
                            meta.setdefault("extra", []).append(v)
                continue
            i += 1
        inicios = [k for k in range(fin) if es_marca(L[k], "ipm:")] + [fin]
        for a, b in zip(inicios, inicios[1:]):
            m = re.search(r"(\d+)", L[a])
            cur = Comp(int(m.group(1)) if m else len(comps) + 1)
            comps.append(cur)
            k = a + 1
            while k < b and (not L[k].strip() or es_marca(L[k], "nombre del componente", "tipo de componente")
                             or re.fullmatch(r"_+", L[k].strip())):
                if es_marca(L[k], "nombre del componente"):
                    cur.titulo = re.sub(r"^.*?[:：]\s*[“\"]?", "", plano_md(L[k])).strip("”\" ")
                elif es_marca(L[k], "tipo de componente"):
                    cur.tipo = re.sub(r"^.*?:\s*", "", plano_md(L[k])).strip(" .")
                k += 1
            cur.items = self.componente(k, b)
        return meta, comps

    def componente(self, i0, i1):
        L = self.L
        items = []
        # cápsula de objetivo
        preg = [k for k in range(i0, i1) if k in self.pregs]
        i = i0
        while i < i1:
            if es_marca(L[i], "ipp: cápsula de objetivo", "ipp: capsula de objetivo"):
                j = i + 1
                lineas = []
                while j < i1 and len(lineas) < 2:
                    s = L[j].strip()
                    if s and (re.match(r"(?i)^recordemos", plano_md(s)) or plano_md(s).startswith("IP")):
                        break
                    if s:
                        lineas.append(sin_negrita(s))
                    j += 1
                items.append(dict(k="objetivo", lineas=lineas))
                i = j
                continue
            if i in preg:
                sig = min([k for k in preg if k > i] + [i1])
                ult = self.ultima_relevante(i, sig)
                n, enun = self.pregs[i]
                claves = []
                cuerpo = self.bloques(i + 1, ult + 1, en_pregunta=True, claves=claves)
                html_c = "".join(self.render(it) for it in cuerpo)
                items.append(dict(k="qmd", n=n, enun="<p class='enun'>%s</p>" % inline(sin_negrita(enun)),
                                  html=html_c, claves=claves))
                i = ult + 1
                continue
            sig = min([k for k in preg if k > i] + [i1])
            j = i
            # tramo hasta la próxima pregunta o la próxima cápsula de objetivo
            while j < sig and not es_marca(L[j], "ipp: cápsula de objetivo", "ipp: capsula de objetivo"):
                j += 1
            if j == i:
                j = i + 1
            items.extend(self.bloques(i, j))
            i = j
        return items

    def ultima_relevante(self, i, sig):
        """Última línea de la pregunta que empieza en i: su clave, su último hueco o alternativa, y las
        notas que siguen pegadas a ella (imágenes que se muestran después de validar)."""
        L = self.L
        ult = i
        for k in range(i + 1, sig):
            s = L[k].strip()
            p = plano_md(s)
            if not p:
                continue
            if (es_hueco(s) or k in self.alts or RE_NUM.match(s) or RE_PARTE.match(s)
                    or tokens_clave(s) is not None and (p.startswith("IPP") or re.match(r"(?i)^(respuestas?|correctas?)\b", p))
                    or (es_tabla(s) and re.search(r"_{3,}|\{.*,.*\}|Completa", s))):
                ult = k
        k = ult + 1
        while k < sig:
            s = L[k].strip()
            p = plano_md(s)
            if not p:
                k += 1
                continue
            if p.startswith(("IPP", "IPD")) and not es_dialogo_ini(s) and not es_capsula(s) and (
                    re.search(r"(?i)validar|comprobar", p) or RE_STATIC.search(s)):
                ult = k
                k += 1
                continue
            break
        return ult

    def render(self, it):
        return render_item(it, "md%d" % id(it), self.personajes, self.orden)


CSS = """
:root{--bg:#fff;--card:#f1f2f4;--opt:#fff;--ink:#1f1b2e;--mut:#6b6880;--pri:#423475;--lin:#dcdde3;--ctx:#fafafb;--acc:#d45762;--soft:#efecf8}
*{box-sizing:border-box}html{scroll-behavior:smooth}
body{margin:0;background:var(--bg);color:var(--ink);font:16px/1.55 system-ui,-apple-system,"Segoe UI",sans-serif}
nav.top{position:sticky;top:0;z-index:20;background:rgba(255,255,255,.94);backdrop-filter:blur(6px);border-bottom:1px solid var(--lin);display:flex;gap:6px;overflow-x:auto;padding:8px 12px}
nav.top a{white-space:nowrap;text-decoration:none;color:var(--pri);font-size:.85rem;padding:5px 12px;border-radius:999px;background:var(--soft)}
nav.top a:hover{background:var(--pri);color:#fff}
header.h{background:linear-gradient(135deg,#423475,#6b5aa8);color:#fff;padding:34px 16px}
header.h .in{max-width:860px;margin:0 auto}header.h small{opacity:.8;letter-spacing:.04em}
header.h h1{margin:6px 0 10px;font-size:1.7rem;line-height:1.25}header.h p{margin:0;opacity:.92}
main{max-width:860px;margin:0 auto;padding:12px 16px 90px}
.comp{margin:44px 0 0;scroll-margin-top:60px}
.comp>.ch{display:flex;align-items:baseline;gap:10px;flex-wrap:wrap;border-bottom:2px solid var(--pri);padding-bottom:8px;margin-bottom:14px}
.comp>.ch h2{margin:0;font-size:1.25rem;color:var(--pri)}.comp>.ch .tag{font-size:.75rem;color:var(--mut);border:1px solid var(--lin);border-radius:999px;padding:2px 10px}
.card{background:var(--card);border:1px solid var(--lin);border-radius:14px;padding:18px 20px;margin:14px 0}
.qn{display:inline-block;font-size:.78rem;font-weight:700;letter-spacing:.05em;text-transform:uppercase;color:var(--pri);margin-bottom:4px}
.enun{font-weight:600;margin:4px 0 8px}
p.txt{margin:12px 0}.card p{margin:8px 0}
img.fig{display:block;max-width:100%;height:auto;margin:10px auto;border-radius:8px;border:1px solid var(--lin);background:#fff}
.figs{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:12px;margin:10px 0}.figs figure{margin:0;text-align:center}
table{border-collapse:collapse;margin:10px 0;width:100%;font-size:.94rem;background:#fff}
.tabw{overflow-x:auto;max-width:100%}
table[data-unalinea] th,table[data-unalinea] td{white-space:nowrap}
th,td{border:1px solid var(--lin);padding:7px 10px;text-align:center;vertical-align:middle}th{background:var(--soft)}
td.izq,th.izq{text-align:left}table.vf td p{margin:0}
input.comp{width:90px;border:1px solid var(--pri);border-radius:6px;padding:4px 6px;font:inherit}
select{font:inherit;border:1px solid var(--pri);border-radius:6px;padding:3px 6px;background:#fff}
ul.l{margin:8px 0;padding-left:22px}
.opts{list-style:none;margin:14px 0 0;padding:0;display:grid;gap:8px}
.opts li{border:1px solid var(--lin);border-radius:10px;padding:10px 14px;background:var(--opt)}
.opts label{display:flex;gap:12px;align-items:flex-start;cursor:pointer}
.opts input{margin-top:5px;accent-color:var(--pri)}
.opts.chk li{display:flex;gap:10px}.opts.chk li::before{content:"";flex:0 0 16px;height:16px;margin-top:4px;border:2px solid var(--pri);border-radius:4px}
.opts.img li{display:block}.opts .lab{font-size:.78rem;color:var(--mut);margin-left:6px}
.caja{background:var(--ctx);border:1px dashed var(--pri);border-radius:14px;padding:16px 20px;margin:14px 0}.caja h4{margin:0 0 6px;color:var(--pri)}
.obj{background:var(--soft);border-left:5px solid var(--pri);border-radius:10px;padding:14px 18px;margin:14px 0}.obj p{margin:6px 0 0}
details.exp{margin:12px 0;border:1px solid var(--lin);border-radius:12px;background:#fff}
details.exp summary{cursor:pointer;padding:10px 16px;font-weight:600;color:var(--pri)}details.exp>div{padding:2px 18px 14px}
details.btn{margin:10px 0}details.btn summary{cursor:pointer;display:inline-block;background:var(--pri);color:#fff;border-radius:8px;padding:6px 14px;font-size:.9rem;list-style:none}
.rec{border:2px dashed var(--acc);border-radius:12px;padding:14px 18px;margin:12px 0;color:var(--mut);font-size:.92rem}.rec span{display:block;font-weight:700;color:var(--acc);font-size:.78rem;text-transform:uppercase;letter-spacing:.05em}
.dia{margin:18px 0;display:grid;gap:12px}
.fila{display:flex;gap:12px;align-items:flex-end}.fila.der{flex-direction:row-reverse}
.avc{flex:0 0 64px;text-align:center}.avc .av{width:64px;height:auto;display:block}.avc .nom{font-size:.72rem;font-weight:700;color:var(--pri)}
.burb{position:relative;background:var(--soft);border-radius:18px;padding:10px 16px;max-width:calc(100% - 90px)}
.fila.izq .burb{border-bottom-left-radius:4px}.fila.der .burb{border-bottom-right-radius:4px;background:#fdecee}
.car{position:relative;margin:18px 0;border:2px solid var(--pri);border-radius:16px;background:#fff;overflow:hidden}
.car-et{background:var(--pri);color:#fff;font-size:.78rem;font-weight:700;letter-spacing:.06em;text-transform:uppercase;padding:6px 16px}
.sub{margin:22px 0 4px;color:var(--pri)}
.sl{display:none;padding:18px 24px 10px;min-height:200px}.sl.on{display:block}.sl h3{margin:0 0 8px;color:var(--pri)}
.sl-n{float:right;font-size:.75rem;color:var(--mut)}.sl-b p{margin:8px 0}
.car-nav{display:flex;align-items:center;justify-content:center;gap:14px;padding:8px 0 12px}
.car-nav button{border:0;background:var(--soft);color:var(--pri);width:36px;height:36px;border-radius:50%;font-size:1.3rem;cursor:pointer}
.car-nav button:hover{background:var(--pri);color:#fff}
.dots i{display:inline-block;width:9px;height:9px;border-radius:50%;background:var(--lin);margin:0 3px}.dots i.on{background:var(--pri)}
.nota{font-size:.82rem;color:var(--mut);margin-top:40px}
input.hz{width:6ch;border:1px solid var(--pri);border-radius:6px;padding:2px 6px;font:inherit;background:#fff;color:var(--ink);margin:0 2px}
select.hz{max-width:100%;margin:0 2px}
.hz.ok{border-color:#2e9e57;background:#e3f4e7}.hz.mal{border-color:#d13b3b;background:#fbe4e4}
.hzbox{display:inline-block}.hzbox input.hz{font-size:.8em;width:5ch;padding:0 4px;margin:0;font-family:system-ui,sans-serif}
.hzbox select.hz{font-size:.75em;font-family:system-ui,sans-serif;margin:0}
.acc{margin-top:16px;display:flex;gap:12px;align-items:center;flex-wrap:wrap}
button.chk{background:var(--pri);color:#fff;border:0;border-radius:8px;padding:7px 18px;font:inherit;font-size:.92rem;cursor:pointer}
button.chk:hover{background:#2f2457}
button.versol{background:#fff;color:var(--pri);border:1px solid var(--pri);border-radius:8px;padding:6px 14px;font:inherit;font-size:.88rem;cursor:pointer}
.res{font-weight:600;font-size:.94rem}.res.ok{color:#2e9e57}.res.mal{color:#d13b3b}.res.neu{color:var(--pri)}
.opts li.ok{background:#e3f4e7;border-color:#2e9e57}.opts li.mal{background:#fbe4e4;border-color:#d13b3b}
.opts li.sol{box-shadow:0 0 0 2px #2e9e57 inset}
tr.ok>td{background:#e3f4e7}tr.mal>td{background:#fbe4e4}tr.sol>td:first-child{box-shadow:4px 0 0 #2e9e57 inset}
.sol{display:inline-block;margin:0 6px;font-size:.8rem;color:#1f7a40;background:#e3f4e7;border-radius:6px;padding:0 6px;font-weight:600}
.tras{display:none!important}
.parte{margin:16px 0 6px}.parte .pl{font-weight:700;color:var(--pri);margin-right:4px}.parte.enun{font-weight:600}
textarea.abierta{width:100%;font:inherit;border:1px solid var(--pri);border-radius:8px;padding:8px 10px;background:#fff}
.ph{border:2px dashed #b9b6c9;border-radius:12px;padding:12px 16px;margin:12px 0;color:var(--mut);font-size:.9rem;background:#fafafb}
.ph span{display:block;font-weight:700;font-size:.75rem;text-transform:uppercase;letter-spacing:.05em;color:#8a86a0}
.recw{margin:14px 0}.recw iframe{display:block;width:100%;height:560px;border:1px solid var(--lin);border-radius:12px;background:#fff}
.recw a{font-size:.8rem;color:var(--pri)}
.car.uno .sl{min-height:0;padding-bottom:16px}
table.selfilas td:first-child,table.selfilas th:first-child{width:44px}
details.exp .cand{display:none;font-weight:400;font-size:.8rem;color:var(--mut);margin-left:10px}
details.exp.bloq{background:#f6f6f8}details.exp.bloq summary{cursor:not-allowed;color:var(--mut)}details.exp.bloq .cand{display:inline}
details.exp:not(.bloq){border-color:var(--pri)}
"""

JS = r"""
document.querySelectorAll('.car').forEach(function(c){
  var s=c.querySelectorAll('.sl'),d=c.querySelectorAll('.dots i'),i=0;
  if(!d.length)return;
  function go(n){i=(n+s.length)%s.length;s.forEach(function(x,k){x.classList.toggle('on',k===i)});d.forEach(function(x,k){x.classList.toggle('on',k===i)});}
  c.querySelector('.prev').onclick=function(){go(i-1)};c.querySelector('.next').onclick=function(){go(i+1)};
  d.forEach(function(x,k){x.onclick=function(){go(k)}});
  c.tabIndex=0;c.addEventListener('keydown',function(e){if(e.key==='ArrowRight'){go(i+1);e.preventDefault();}if(e.key==='ArrowLeft'){go(i-1);e.preventDefault();}});
});
/* ---- validación de respuestas ---- */
function b64d(s){try{return decodeURIComponent(escape(atob(s)))}catch(e){return ''}}
function nrm(s){s=(s||'').normalize('NFD').replace(/[̀-ͯ]/g,'').toLowerCase();
  s=s.replace(/\\\(|\\\)|\$|`/g,'').replace(/\\d?frac\{([^{}]*)\}\{([^{}]*)\}/g,'$1/$2').replace(/\\(?:text|mathrm)\{([^{}]*)\}/g,'$1')
     .replace(/\\cdot|\\times|[·×*]/g,'·').replace(/\\div/g,'÷').replace(/[\s{}()\[\],\\]/g,'').replace(/[.;:]+$/,'');return s;}
function num(s){var t=(s||'').replace(/\s/g,'').replace(',','.');return /^-?\d+(\.\d+)?$/.test(t)?parseFloat(t):null;}
function igual(a,k){var x=num(a),y=num(k);if(x!==null&&y!==null)return Math.abs(x-y)<1e-9;return nrm(a)!==''&&nrm(a)===nrm(k);}
function marca(el,ok){el.classList.remove('ok','mal');el.classList.add(ok?'ok':'mal');}
function evaluar(c){
  var k=b64d(c.dataset.k),t=c.dataset.tipo;
  if(t==='input'){marca(c,igual(c.value,k));return c.classList.contains('ok');}
  if(t==='select'){var ok=c.value!==''&&nrm(c.options[c.selectedIndex].text)===nrm(k);marca(c,ok);return ok;}
  if(t==='radio'){var s=c.querySelector('input:checked');c.querySelectorAll('li').forEach(function(l){l.classList.remove('ok','mal')});
    if(!s)return false;var li=s.closest('li'),ok=li.dataset.i===k;marca(li,ok);return ok;}
  if(t==='vf'){var s2=c.querySelector('input:checked'),ok2=!!s2&&s2.value===k;marca(c,ok2);return ok2;}
  if(t==='multi'){var cl=k.split(',').map(function(x){return x.trim()}),bien=true;
    if(c.tagName==='TABLE'){c.querySelectorAll('tr').forEach(function(tr){var cb=tr.querySelector('input[type=checkbox]');if(!cb)return;
        tr.classList.remove('ok','mal');var dentro=cl.indexOf(cb.value)>=0;if(cb.checked)marca(tr,dentro);if(cb.checked!==dentro)bien=false;});return bien;}
    c.querySelectorAll('li').forEach(function(l){var cb=l.querySelector('input');l.classList.remove('ok','mal');var dentro=cl.indexOf(l.dataset.i)>=0;
        if(cb.checked)marca(l,dentro);if(cb.checked!==dentro)bien=false;});return bien;}
  return true;
}
function mostrarSol(c){
  var k=b64d(c.dataset.k),t=c.dataset.tipo;
  if(t==='input'||t==='select'){if(!(c.nextElementSibling&&c.nextElementSibling.classList.contains('sol'))){var s=document.createElement('span');s.className='sol';s.textContent=k;c.insertAdjacentElement('afterend',s);}return;}
  if(t==='radio'){var li=c.querySelector("li[data-i='"+k+"']");if(li)li.classList.add('sol');return;}
  if(t==='vf'){var td=c.lastElementChild;if(!td.querySelector('.sol')){var s2=document.createElement('span');s2.className='sol';s2.textContent=k;c.firstElementChild.appendChild(s2);}return;}
  if(t==='multi'){var cl=k.split(',');if(c.tagName==='TABLE'){c.querySelectorAll('tr').forEach(function(tr){var cb=tr.querySelector('input[type=checkbox]');if(cb&&cl.indexOf(cb.value)>=0)tr.classList.add('sol');});}
    else c.querySelectorAll('li').forEach(function(l){if(cl.indexOf(l.dataset.i)>=0)l.classList.add('sol');});}
}
function revelar(card){
  card.querySelectorAll('.tras').forEach(function(x){x.classList.remove('tras')});
  card.querySelectorAll('details.exp.bloq').forEach(function(x){x.classList.remove('bloq');x.open=true;});
  var e=card.nextElementSibling;
  while(e&&!(e.classList.contains('card')&&e.classList.contains('q'))){
    if(e.classList.contains('tras'))e.classList.remove('tras');
    e.querySelectorAll('.tras').forEach(function(x){x.classList.remove('tras')});
    if(e.matches('details.exp.bloq')){e.classList.remove('bloq');e.open=true;}
    e=e.nextElementSibling;}
}
document.querySelectorAll('.card.q').forEach(function(card){
  var btn=card.querySelector('.acc .chk'),res=card.querySelector('.acc .res'),vs=card.querySelector('.acc .versol');
  if(!btn)return;
  var ctl=[].slice.call(card.querySelectorAll('[data-k]')).filter(function(c){return c.closest('.card.q')===card});
  btn.addEventListener('click',function(){
    if(!ctl.length){res.textContent=btn.textContent==='Enviar'?'Respuesta enviada.':'';res.className='res neu';revelar(card);return;}
    var ok=0;ctl.forEach(function(c){if(evaluar(c))ok++;});
    if(ok===ctl.length){res.textContent='¡Correcto!';res.className='res ok';}
    else{res.textContent='Correctas: '+ok+' de '+ctl.length+'. Revisa las marcadas en rojo.';res.className='res mal';}
    vs.hidden=false;revelar(card);
  });
  vs.addEventListener('click',function(){ctl.forEach(mostrarSol);});
  card.addEventListener('input',function(e){var el=e.target;if(el.matches('input.hz'))el.style.width=Math.max(6,el.value.length+2)+'ch';});
  card.addEventListener('change',function(e){var el=e.target.closest('[data-k],li,tr');if(el)el.classList.remove('ok','mal');res.textContent='';});
});
document.querySelectorAll('details.exp').forEach(function(d){d.addEventListener('toggle',function(){if(d.classList.contains('bloq')&&d.open)d.open=false;});});
/* un Exploremos sin pregunta antes (en su sección) no tiene qué esperar */
document.querySelectorAll('details.exp.bloq').forEach(function(d){if(d.closest('.card.q'))return;var e=d.previousElementSibling;
  while(e&&!(e.classList.contains('card')&&e.classList.contains('q')))e=e.previousElementSibling;if(!e)d.classList.remove('bloq');});
/* huecos dentro de fórmulas: KaTeX dibuja una caja con id y el control se mueve adentro */
function colocarHuecos(){document.querySelectorAll('.pend').forEach(function(p){var box=document.getElementById(p.dataset.hz);
  if(box&&p.firstElementChild){box.innerHTML='';box.classList.add('hzbox');box.appendChild(p.firstElementChild);p.remove();}});}
function dibujarMate(){renderMathInElement(document.body,{delimiters:[{left:'\\(',right:'\\)',display:false},{left:'\\[',right:'\\]',display:true}],
  throwOnError:false,strict:'ignore',trust:function(c){return c.command==='\\htmlId'}});colocarHuecos();}
"""


def meta_header(meta):
    titulo = obj = ""
    extra = []
    for t in meta.get("tablas", []):
        for f in t:
            cel = [texto_celda(c) for c in f]
            if len(cel) >= 2:
                k, v = cel[0], cel[1]
                if k.startswith("Título"):
                    titulo = v
                elif k.startswith("Objetivo"):
                    obj = v
                elif k in ("Unidad", "Tema", "Curso"):
                    extra.append(v)
    return titulo, obj, extra


def generar(doc, titulo_forzado=None):
    _CTL.clear()
    _hz[0] = 0
    if isinstance(doc, LectorMd):
        doc.personajes = doc.orden = personajes_md(doc.L)
        meta, comps = doc.leer()
        personajes = doc.personajes
        titulo, obj, extra = meta.get("titulo", ""), meta.get("obj", ""), meta.get("extra", [])
    else:
        meta, comps, personajes = parsear(doc)
        titulo, obj, extra = meta_header(meta)
    titulo = titulo_forzado or titulo or "Lección"
    orden = list(personajes)
    nav = "".join("<a href='#c%d'>%d · %s</a>" % (c.n, c.n, html.escape(re.sub(r'^Lecci[oó]n \d+:\s*', '', c.titulo) or "Componente")) for c in comps)
    h = ["<!doctype html><html lang='es'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'>",
         "<title>%s</title>" % html.escape(titulo),
         "<link rel='stylesheet' href='https://cdnjs.cloudflare.com/ajax/libs/KaTeX/0.16.9/katex.min.css'><style>%s</style></head><body>" % CSS,
         "<nav class='top'>%s</nav>" % nav,
         "<header class='h'><div class='in'><small>%s</small><h1>%s</h1><p>%s</p></div></header><main>" % (
             html.escape(" · ".join(extra[1:])), html.escape(titulo), html.escape(obj))]
    cid = 0
    for c in comps:
        h.append("<section class='comp' id='c%d'><div class='ch'><h2>%s</h2><span class='tag'>%s</span></div>" % (
            c.n, html.escape(re.sub(r'^Lecci[oó]n \d+:\s*', '', c.titulo) or "Componente %d" % c.n), html.escape(c.tipo)))
        # Un recuadro de "recurso por producir" (IPC) sobra si el recurso ya viene incrustado enseguida.
        vis = [it for i, it in enumerate(c.items)
               if not (it["k"] == "recurso" and any("<iframe" in (x.get("html") or "")
                                                    for x in c.items[i + 1:i + 4]))]
        for it in vis:
            cid += 1
            h.append(render_item(it, "car%d" % cid, personajes, orden))
        h.append("</section>")
    h.append("</main><script>%s</script>" % JS)
    h.append("<script defer src='https://cdnjs.cloudflare.com/ajax/libs/KaTeX/0.16.9/katex.min.js'></script>"
             "<script defer src='https://cdnjs.cloudflare.com/ajax/libs/KaTeX/0.16.9/contrib/auto-render.min.js' "
             "onload='dibujarMate()'></script></body></html>")
    salida = re.sub(r" data-k='\x01\d+\x01'", "", "".join(h))       # controles sin clave
    return salida, comps


def personajes_md(L):
    """Personajes de los diálogos del .md, en orden de aparición (el primero va a la izquierda)."""
    vistos, dentro = [], False
    for l in L:
        if es_dialogo_ini(l):
            dentro = True
            continue
        if es_dialogo_fin(l):
            dentro = False
            continue
        if dentro:
            m = re.match(r"^\**\s*([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+)\s*\**\s*:", l.strip())
            if m and m.group(1) not in vistos and m.group(1) not in NO_PERSONAJES:
                vistos.append(m.group(1))
    return vistos


def docx_de_md(ruta):
    """True si el .docx lo produjo guion_a_docx.py (sin listas de Word) y su .md está al lado."""
    md = os.path.splitext(ruta)[0] + ".md"
    if not os.path.exists(md):
        return None
    try:
        d = docx.Document(ruta)
    except Exception:
        return None
    if d.element.body.find(".//" + qn("w:numPr")) is not None:
        return None
    return md


def main():
    ap = argparse.ArgumentParser(description="Guion de lección completa (.docx del Doc o .md) -> HTML de lectura.")
    ap.add_argument("entrada")
    ap.add_argument("-o", "--salida")
    ap.add_argument("--titulo")
    a = ap.parse_args()
    entrada = a.entrada
    if entrada.lower().endswith(".docx"):
        md = docx_de_md(entrada)
        if md:
            print("Este .docx salió de guion_a_docx.py; se lee su .md:", md, file=sys.stderr)
            entrada = md
    doc = LectorMd(entrada) if entrada.lower().endswith(".md") else Doc(entrada)
    salida = a.salida or os.path.splitext(a.entrada)[0] + "_vista.html"
    h, comps = generar(doc, a.titulo)
    open(salida, "w", encoding="utf-8").write(h)
    nq = sum(1 for c in comps for it in c.items if it["k"] in ("q", "qmd"))
    print(salida, "·", len(comps), "componentes ·", nq, "preguntas ·", len(h) // 1024, "KB")


if __name__ == "__main__":
    main()
