"""Baja un Google Doc a .md con el formato del proyecto.

Entrada: un .docx (bajado a mano desde Google) o el JSON del conector de Drive
con la exportación .docx en base64.

    python3 doc_a_md.py <archivo.docx|tool-result.json> <destino.md>

Pipeline: .docx -> preparar XML -> pandoc gfm -> desarmado de cajas -> normalización -> .md

Las listas de Docs se resuelven en el XML, porque pandoc no conserva su formato: las
decimales en negrita son enunciados (`**N.** texto`), las de letras son alternativas
(sin letra) o, si todos sus ítems van en negrita, partes (`**a. texto**`).

Las tablas de UNA columna del Doc son cajas (cápsulas, diálogos, cuadros de
respuesta): se despliegan a prosa con separador, que es el formato del proyecto.
Las multicolumna son datos y se conservan como tablas pipe.

El HTML que emite pandoc para celdas de varios párrafos se convierte a mano, sin
volver a pasar por pandoc: el round-trip escapaba el LaTeX (\\( \\Omega \\)).
"""
import json, io, base64, re, subprocess, sys, tempfile, os, html as _html
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import alternativas  # noqa: E402

def celda_a_md(frag):
    t = frag
    # Negritas vacías y pegadas se resuelven aquí: ya en markdown un "**\n\n**" entre dos
    # párrafos en negrita no se distingue de una negrita vacía, y se fundían los párrafos.
    # Un salto dentro de la negrita la parte: el ** no puede cruzar líneas.
    def sin_saltos(m):
        trozos = re.split(r'<br\s*/?>\s*', m.group(2))
        return "<br />".join("<strong>%s</strong>" % x if x.strip() else "" for x in trozos)
    t = re.sub(r'<(strong|b)>(.*?)</\1>', sin_saltos, t, flags=re.S)
    t = re.sub(r'</strong>[ \t]*<strong>', '', t)
    t = re.sub(r'<br\s*/?>', '\n', t)
    t = re.sub(r'<li[^>]*>\s*(?:<p[^>]*>)?', '\n\n- ', t)   # viñeta y texto en la misma línea
    t = re.sub(r'</p>\s*<p[^>]*>', '\n\n', t)
    t = re.sub(r'</?p[^>]*>', '\n\n', t)
    t = re.sub(r'</?(strong|b)>', '**', t)
    t = re.sub(r'</?(em|i)>', '*', t)
    t = re.sub(r'<img[^>]*src="[^"]*?([^/"]+)"[^>]*/?>', r'**[imagen: \1]**', t)
    t = re.sub(r'<[^>]+>', '', t)
    t = _html.unescape(t)
    t = t.replace('****', '')                    # imagen entre negritas contiguas
    t = re.sub(r'[ \t]+', ' ', t)
    t = re.sub(r'\n{3,}', '\n\n', t)
    return "\n\n".join(p.strip() for p in t.split("\n\n") if p.strip())

def procesar_tablas_html(t):
    def rep(m):
        tabla = m.group(0)
        filas = re.findall(r'<tr[^>]*>(.*?)</tr>', tabla, re.S)
        filas = [re.findall(r'<t[dh][^>]*>(.*?)</t[dh]>', f, re.S) for f in filas]
        filas = [f for f in filas if f]
        if not filas: return ""
        ncol = max(len(f) for f in filas)
        if ncol == 1:                                    # caja -> prosa
            partes = [celda_a_md(c) for f in filas for c in f]
            return "\n\n" + "\n\n".join(p for p in partes if p) + "\n\n---\n"
        # datos -> tabla pipe (párrafos unidos en una línea)
        out = []
        for f in filas:
            cs = [re.sub(r'\s*\n+\s*', ' ', celda_a_md(c)).strip() for c in f]
            cs += [""] * (ncol - len(cs))
            out.append("| " + " | ".join(cs) + " |")
        out.insert(1, "|" + "----|" * ncol)
        return "\n\n" + "\n".join(out) + "\n"
    # Las tablas pueden venir anidadas (una tabla dentro de una cápsula).
    # Se procesan de adentro hacia afuera: el patrón exige que no haya otro
    # <table> dentro, y se repite hasta que no queda ninguna.
    interna = r'<table(?:(?!<table).)*?</table>'
    for _ in range(10):
        t, n = re.subn(interna, rep, t, flags=re.S)
        if not n: break
    return t

def desarmar_cajas_pipe(t):
    """Cajas que quedaron como tabla pipe de una sola columna."""
    ls, out, i = t.split("\n"), [], 0
    while i < len(ls):
        if (i + 1 < len(ls) and re.fullmatch(r'\|\s*\|', ls[i].strip())
                and re.fullmatch(r'\|[-: ]+\|', ls[i+1].strip())):
            i += 2
            cuerpo = []
            while i < len(ls) and ls[i].strip().startswith("|"):
                cuerpo.append(re.sub(r'\|$', '', re.sub(r'^\|', '', ls[i].strip())).strip())
                i += 1
            if out and out[-1].strip(): out.append("")
            for p in cuerpo:
                if p: out.append(p); out.append("")
            out.append("---"); out.append("")
            continue
        out.append(ls[i]); i += 1
    return "\n".join(out)

def limpiar_html_suelto(t):
    """Etiquetas que pandoc deja fuera de las tablas."""
    t = re.sub(r'<img[^>]*src="[^"]*?([^/"]+)"[^>]*/?>', r'**[imagen: \1]**', t)
    t = re.sub(r'</?u>', '', t)
    t = re.sub(r'</?(span|font)[^>]*>', '', t)
    return t

def cerrar_negritas(t):
    """En el Doc, un salto de línea dentro de una celda deja el ** abierto:
    '**Título\n**Cuerpo'. Se cierra el título y se separa en dos párrafos."""
    # El ** de apertura va al inicio de línea: si no, el cierre de "**Diego:** texto" se toma
    # por apertura y se corren las negritas de todo el diálogo.
    t = re.sub(r'^\*\*([^*\n]+?)\n\s*\n?\*\*(?=\S)', r'**\1**\n\n', t, flags=re.M)
    t = re.sub(r'^\*\*([^*\n]+?)\n\s*\n\*\*\s*\n', r'**\1**\n\n', t, flags=re.M)  # cierre en línea sola
    t = re.sub(r'\*\*([^*\n]+?)\*\*\*\*', r'**\1**', t)   # ** duplicado al cerrar
    return t

def desescapar(t):
    for a, b in [(r'\_', '_'), (r'\.', '.'), (r'\!', '!'), (r'\<', '<'),
                 (r'\>', '>'), (r'\#', '#'), (r'\$', '$'), (r'\|', '|')]:
        t = t.replace(a, b)
    return t

NOMBRE = r"[A-ZÁÉÍÓÚÑ]\w*(?: [A-ZÁÉÍÓÚÑ]\w*){0,2}"
CLAVE_LETRA = re.compile(
    r"^IPP:\s*(?:alternativa\s+|respuesta\s+)?correctas?\s*:?\s*\{?\s*([a-e](?:\s*[,;y]\s*[a-e])*)\s*\}?\s*(?:\.\s*(.*))?$",
    re.I)
LISTA_XML = re.compile(r'<w:tbl(?:\s[^>]*)?>|</w:tbl>|<w:p(?:\s[^>]*[^/>])?>.*?</w:p>', re.S)


def formatos_de_lista(zi):
    """numId -> (id de lista, inicio forzado o None, {nivel: (formato, inicio)})."""
    import xml.etree.ElementTree as ET
    if "word/numbering.xml" not in zi.namelist():
        return {}
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    root = ET.fromstring(zi.read("word/numbering.xml"))
    abstractos = {}
    for a in root.findall(W + "abstractNum"):
        niveles = {}
        for l in a.findall(W + "lvl"):
            fmt, ini = l.find(W + "numFmt"), l.find(W + "start")
            niveles[l.get(W + "ilvl")] = (fmt.get(W + "val") if fmt is not None else "",
                                          int(ini.get(W + "val")) if ini is not None else 1)
        abstractos[a.get(W + "abstractNumId")] = niveles
    out = {}
    for n in root.findall(W + "num"):
        ref = n.find(W + "abstractNumId").get(W + "val")
        o = n.find(W + "lvlOverride/" + W + "startOverride")
        out[n.get(W + "numId")] = (ref, int(o.get(W + "val")) if o is not None else None, abstractos.get(ref, {}))
    return out


def fraccion_negrita(p):
    total = neg = 0
    for r in re.findall(r'<w:r(?:\s[^>]*)?>.*?</w:r>', p, re.S):
        largo = len("".join(re.findall(r'<w:t(?:\s[^>]*)?>([^<]*)</w:t>', r)).strip())
        total += largo
        m = re.search(r'<w:b(?:\s+w:val="([^"]*)")?\s*/>', r)
        if m and m.group(1) not in ("0", "false"):
            neg += largo
    return neg / total if total else 0


def marcar_listas(x, formatos):
    """Resuelve las listas fuera de tablas: pandoc emite "1." para cualquier formato.
    Deja @@N<n>@@ en los enunciados, @@P<letra>@@ en las partes, y quita la marca de lista de
    las alternativas. Las decimales que no van en negrita quedan como están."""
    def lista_de(p):
        m = re.search(r'<w:numPr>(.*?)</w:numPr>', p, re.S)
        if not m: return None
        i = re.search(r'<w:numId w:val="(\d+)"', m.group(1))
        if not i or i.group(1) not in formatos: return None
        n = re.search(r'<w:ilvl w:val="(\d+)"', m.group(1))
        return i.group(1), n.group(1) if n else "0"
    stats, prof = {}, 0
    for m in LISTA_XML.finditer(x):
        s = m.group(0)
        if s.startswith("<w:tbl"): prof += 1
        elif s == "</w:tbl>": prof -= 1
        elif prof == 0:
            l = lista_de(s)
            if l and re.search(r'<w:t(?:\s[^>]*)?>[^<]*\S', s):
                stats.setdefault(l, []).append(fraccion_negrita(s) >= 0.8)
    cuentas, vistos = {}, set()
    prof = 0
    def sub(m):
        nonlocal prof
        s = m.group(0)
        if s.startswith("<w:tbl"): prof += 1; return s
        if s == "</w:tbl>": prof -= 1; return s
        if prof: return s
        l = lista_de(s)
        if not l: return s
        num, nivel = l
        ref, forzado, niveles = formatos[num]
        fmt, ini = niveles.get(nivel, ("", 1))
        if num not in vistos:
            vistos.add(num)
            if forzado is not None: cuentas[(ref, nivel)] = forzado - 1
        n = cuentas.get((ref, nivel), ini - 1) + 1
        cuentas[(ref, nivel)] = n
        for clave in [c for c in cuentas if c[0] == ref and int(c[1]) > int(nivel)]:
            del cuentas[clave]
        marcas = stats.get(l, [])
        if not marcas: marca = ""                                  # ítem vacío
        elif fmt == "decimal":
            if nivel != "0" or sum(marcas) * 2 <= len(marcas): return s
            marca = "@@N%d@@" % n
        elif fmt in ("lowerLetter", "upperLetter"):
            marca = "@@P%s@@" % chr(96 + n) if all(marcas) and n <= 26 else ""
        else:
            return s
        s = re.sub(r'<w:numPr>.*?</w:numPr>', '', s, count=1, flags=re.S)
        if marca:
            run = '<w:r><w:t xml:space="preserve">%s</w:t></w:r>' % marca
            if "</w:pPr>" in s: s = s.replace("</w:pPr>", "</w:pPr>" + run, 1)
            else: s = re.sub(r'^(<w:p(?:\s[^>]*)?>)', lambda g: g.group(1) + run, s, count=1)
        return s
    return LISTA_XML.sub(sub, x)


def preparar_xml(x, formatos):
    # pandoc descarta columnas cuando un ancho de w:gridCol no es entero, y Docs los exporta
    # así ("1100.9999999999995"): por eso se perdían celdas enteras, como las X de un V/F.
    x = re.sub(r'(<w:gridCol w:w=")(\d+\.\d+)(")',
               lambda m: m.group(1) + str(round(float(m.group(2)))) + m.group(3), x)
    x = marcar_listas(x, formatos)
    # pandoc >= 3.x altera el LaTeX (\\( \\) y \\, \\%): se protegen las barras invertidas antes y se restauran después.
    return x.replace("\\", "@@B@@")


def quitar_negrita_inicial(s):
    """'**Texto:** ___' -> 'Texto: ___'; deja igual lo demás."""
    m = re.match(r'^\*\*((?:(?!\*\*).)+)\*\*(.*)$', s)
    if m and "**" not in m.group(2):
        return m.group(1) + m.group(2)
    return s


def quitar_letras_alternativas(t):
    """Alternativas con la letra escrita a mano en el Doc: una tanda a., b., c. seguida de su
    clave, o de una marca IPx si la tanda sigue a un enunciado que termina en pregunta."""
    ls = t.split("\n")
    i = 0
    while i < len(ls):
        if not re.match(r'^a\. \S', ls[i]):
            i += 1
            continue
        tanda, j, esperada = [], i, 0
        while j < len(ls):
            if not ls[j].strip():
                j += 1
                continue
            if ls[j].startswith("abcde"[esperada:esperada + 1] + ". ") and esperada < 5:
                tanda.append(j)
                esperada += 1
                j += 1
            else:
                break
        previo = next((ls[k].strip() for k in range(i - 1, -1, -1) if ls[k].strip()), "")
        sigue = ls[j].strip() if j < len(ls) else ""
        clave = re.match(r'^IPP: Correctas? \{', sigue) or CLAVE_LETRA.match(sigue)
        tras_pregunta = previo.replace("*", "").endswith("?") and sigue.startswith(("IPP", "IPC", "IPD", "IPM"))
        if len(tanda) >= 2 and (clave or tras_pregunta):
            for k in tanda:
                ls[k] = ls[k][3:]
        i = max(j, i + 1)
    return "\n".join(ls)


def tablas_con_encabezado(t):
    """Encabezado vacío y primera fila en negrita: esa fila sube a encabezado."""
    def fila(m):
        filas = m.group(0).rstrip("\n").split("\n")
        if (len(filas) >= 3 and not re.sub(r'[|\s]', '', filas[0]) and re.fullmatch(r'[|:\-\s]+', filas[1])
                and "**" in filas[2]):
            filas = [filas[2], "|" + "---|" * (filas[2].count("|") - 1)] + filas[3:]
        return "\n".join(filas) + "\n"
    return re.sub(r'(?:^\|.*\|\n)+', fila, t, flags=re.M)


def parte_en_negrita(letra, texto):
    """Una parte es enunciado, así que va completa en negrita; el espacio de respuesta queda fuera."""
    texto = texto.replace("**", "").strip()
    cola = re.match(r'^(.*?)((?:[ \t]*_{3,}(?:[ \t]*\{[^{}]*\})?)+)[ \t]*$', texto)
    if cola and cola.group(1).strip():
        return "**%s. %s** %s" % (letra, cola.group(1).rstrip(), cola.group(2).strip())
    return "**%s. %s**" % (letra, texto)


def normalizar(t):
    """Del markdown de pandoc al formato de los guiones del proyecto."""
    t = re.sub(r'^<!-- -->[ \t]*$', '', t, flags=re.M)
    t = re.sub(r'\[(https?://[^\]\s]+)\]\(\1\)', r'\1', t)
    # Saltos duros: "\**" al final de un enunciado, y "\" + línea, que pasa a párrafo.
    t = re.sub(r'(?<!\\)\\\*\*[ \t]*$', '**', t, flags=re.M)
    t = re.sub(r'(?<!\\)\\\n', '\n\n', t)
    # Una negrita que un salto de línea dejó abierta entre una línea y el nombre que sigue.
    t = re.sub(r'^\*\*([^*\n]+)\n\s*\n(%s)(?:\*\*:|:\*\*)[ \t]*' % NOMBRE, r'**\1**\n\n**\2:** ', t, flags=re.M)
    t = re.sub(r'\*\*\n\s*\n(%s)(?:\*\*:|:\*\*)[ \t]*' % NOMBRE, r'\n\n**\1:** ', t)
    # Personajes: "**Nombre**: texto" a "**Nombre:** texto", si el nombre interviene más de una vez.
    nombres = re.findall(r'^\*\*(%s)(?:\*\*:|:\*\*) ' % NOMBRE, t, flags=re.M)
    def personaje(m):
        return "**%s:** " % m.group(1) if nombres.count(m.group(1)) > 1 else m.group(0)
    t = re.sub(r'^\*\*(%s)\*\*: ' % NOMBRE, personaje, t, flags=re.M)
    # Diapositivas
    t = re.sub(r'^\*\*Diapositiva (\d+)\*\*[ \t]*$', r'*Diapositiva \1*', t, flags=re.M)
    t = re.sub(r'^\*\*Diapositiva (\d+)\s*(?=[^\s*\d])([^*\n]+?)\*\*(.*)$',
               lambda m: "*Diapositiva %s*\n\n**%s**%s" % (m.group(1), m.group(2).strip(),
                                                         "\n\n" + m.group(3).lstrip(": ") if m.group(3).strip() else ""),
               t, flags=re.M)
    # Marcas IPx en negrita
    t = re.sub(r'^\*\*(IP[PCDM]:[^*\n]*)\*\*[ \t]*$', r'\1', t, flags=re.M)
    # Título del Exploremos: siempre igual, en negrita (guion_a_docx lo lleva a su caja)
    t = re.sub(r'^(?:\*\*)?Exploremos una posible respuesta[.:]?(?:\*\*)?[ \t]*$', '**Exploremos una posible respuesta**', t, flags=re.M)
    # Enunciados y partes
    t = re.sub(r'^@@N(\d+)@@[ \t]*(.*)$', lambda m: "**%s.** %s" % (m.group(1), quitar_negrita_inicial(m.group(2))), t, flags=re.M)
    t = re.sub(r'^@@P([a-z])@@[ \t]*(.*)$', lambda m: parte_en_negrita(m.group(1), quitar_negrita_inicial(m.group(2))), t, flags=re.M)
    t = re.sub(r'^\*\*(\d{1,2})\. ([^\n]*?)\*\*[ \t]*$', r'**\1.** \2', t, flags=re.M)
    # Claves: 'IPP: correcta b.' a 'IPP: Correcta {b}', y lo que venga después en su propia línea.
    def clave(m):
        letras = [c.lower() for c in re.findall(r'[a-e]', m.group(1), re.I)]
        out = "IPP: %s {%s}" % ("Correctas" if len(letras) > 1 else "Correcta", ", ".join(letras))
        return out + ("\n\nIPP: " + m.group(2) if m.group(2) else "")
    t = "\n".join(CLAVE_LETRA.sub(clave, l) for l in t.split("\n"))
    t = quitar_letras_alternativas(t)
    return tablas_con_encabezado(t)


def convertir(src, dst):
    """src puede ser un .docx bajado a mano desde Google (Archivo > Descargar >
    Microsoft Word), o el JSON del conector de Drive con el .docx en base64."""
    borrar = None
    if src.lower().endswith(".docx"):
        docx = src
    else:
        d = json.load(io.open(src, encoding="utf-8"))
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            f.write(base64.b64decode(d["content"]))
        docx = borrar = f.name
    import zipfile
    prot = tempfile.NamedTemporaryFile(suffix=".docx", delete=False); prot.close()
    with zipfile.ZipFile(docx) as zi, zipfile.ZipFile(prot.name, "w", zipfile.ZIP_DEFLATED) as zo:
        formatos = formatos_de_lista(zi)
        for it in zi.infolist():
            d = zi.read(it.filename)
            if it.filename == "word/document.xml":
                d = preparar_xml(d.decode("utf-8"), formatos).encode("utf-8")
            zo.writestr(it, d)
    md = subprocess.run(["pandoc", prot.name, "-t", "gfm", "--wrap=none"],
                        capture_output=True, text=True, check=True).stdout.replace("@@B@@", "\\")
    os.unlink(prot.name)
    if borrar: os.unlink(borrar)
    md = procesar_tablas_html(md)
    md = desarmar_cajas_pipe(md)
    md = limpiar_html_suelto(md)
    md = re.sub(r'^> ?', '', md, flags=re.M)     # sangría del Doc que pandoc lee como cita
    md = cerrar_negritas(md)
    md = desescapar(md)
    md = normalizar(md)
    md = re.sub(r'\n{3,}', '\n\n', md)
    md = re.sub(r'(?:\n---\n\n?){2,}', '\n---\n\n', md)
    md = re.sub(r'[ \t]+\n', '\n', md)
    md = reparar_cabecera(md, src)
    # Convención: en el .md las alternativas van sin numeración ni letras (en el Doc, con letra).
    md = alternativas.limpiar_md(md)
    io.open(dst, "w", encoding="utf-8").write(md.strip() + "\n")
    return md

def tablas_del_docx(docx):
    """Lee las tablas directamente del XML del .docx, sin pasar por pandoc."""
    import zipfile, xml.etree.ElementTree as ET
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    root = ET.fromstring(zipfile.ZipFile(docx).read("word/document.xml"))
    out = []
    for tbl in root.iter(W + "tbl"):
        filas = []
        for tr in tbl.findall(W + "tr"):
            celdas = []
            for tc in tr.findall(W + "tc"):
                ps = [t for t in ("".join(x.text or "" for x in pp.iter(W + "t")).strip()
                                  for pp in tc.iter(W + "p")) if t]
                celdas.append(ps)
            if celdas: filas.append(celdas)
        if filas: out.append(filas)
    return out


def reparar_cabecera(md, docx):
    """pandoc descarta la segunda columna de la tabla IPD: Personajes / IPP.
    Se reconstruye leyendo esa tabla del XML."""
    if not docx.lower().endswith(".docx"): return md
    tabla = None
    for t in tablas_del_docx(docx):
        if any("IPD: Personajes" in p for f in t for c in f for p in c):
            tabla = t; break
    if not tabla: return md
    filas = []
    for f in tabla:
        cs = [" ".join(c).strip() for c in f]
        cs += [""] * (2 - len(cs))
        filas.append("| " + " | ".join(cs[:2]) + " |")
    nueva = "| | |\n| :-: | :-: |\n" + "\n".join(filas)
    roto = re.compile(r'\n+IPD: Personajes\n+IPP\n+---\n', re.M)
    if roto.search(md):
        return roto.sub("\n\n" + nueva + "\n\n", md, count=1)
    return md


def verificar(docx, dst):
    """Compara el texto del .docx contra el .md generado y avisa de lo que falte.

    pandoc puede descartar celdas de tablas que no logra representar (visto en
    la segunda columna de la tabla de Personajes). Este chequeo lo detecta."""
    import zipfile
    if not docx.lower().endswith(".docx"):
        return []
    import xml.etree.ElementTree as ET
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    root = ET.fromstring(zipfile.ZipFile(docx).read("word/document.xml"))
    # Word parte las palabras en varios runs: se unen dentro del párrafo, no entre ellos.
    def texto_p(p):
        partes = []
        for e in p.iter():
            if e.tag == W + "t": partes.append(e.text or "")
            elif e.tag in (W + "br", W + "tab"): partes.append(" ")
        return "".join(partes)
    xml = " ".join(texto_p(p) for p in root.iter(W + "p"))
    md = io.open(dst, encoding="utf-8").read()
    # Word parte y pega palabras entre runs y saltos, así que se compara también
    # contra el .md sin espacios: evita falsos positivos como "ecibir" o "torneoesta".
    plano = re.sub(r'\s+', '', md.lower())
    faltan = []
    for w in sorted(set(re.findall(r'\w{6,}', xml.lower()))):
        if set(w) <= set("_"): continue
        if w in plano: continue
        faltan.append(w)
    return faltan


def verificar_estructura(docx, dst):
    """Detecta lo que `verificar` no ve: párrafos cortos y celdas perdidas. Las palabras de seis
    letras no notan que falta una X de un V/F. Devuelve (párrafos, filas) sin reflejo en el .md."""
    import xml.etree.ElementTree as ET, zipfile
    if not docx.lower().endswith(".docx"):
        return [], []
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    solo = lambda s: re.sub(r'[\W_]+', '', s).lower()
    with io.open(dst, encoding="utf-8") as f:
        md = f.read()
    # Lo que el .md agrega al texto: imágenes y destino de los enlaces.
    md = re.sub(r'\[imagen: [^\]]*\]', '', md)
    md = re.sub(r'\[([^\]]*)\]\([^)\s]*\)', r'\1', md)
    plano = solo(md)
    # La clave se reescribe y la letra escrita a mano se quita: no cuentan como pérdida.
    root = ET.fromstring(zipfile.ZipFile(docx).read("word/document.xml"))
    parrafos = []
    for p in root.iter(W + "p"):
        # Cada línea por separado: un salto puede dejar varias alternativas en un párrafo.
        lineas = [""]
        for e in p.iter():
            if e.tag == W + "t": lineas[-1] += e.text or ""
            elif e.tag == W + "br": lineas.append("")
        for s in lineas:
            s = s.strip()
            if not s or CLAVE_LETRA.match(s) or re.match(r'^IPP:\s*(respuesta\s+)?correcta', s, re.I):
                continue
            n = solo(re.sub(r'^(?:[a-e]|\d{1,2})[.)]\s+', '', s))
            if n and n not in plano:
                parrafos.append(s[:70])
    filas_md = [solo("".join(c)) for c in
                ([x for x in l.strip().strip("|").split("|")] for l in md.split("\n") if l.strip().startswith("|"))]
    filas = []
    for t, tabla in enumerate(tablas_del_docx(docx), 1):
        for i, f in enumerate(tabla, 1):
            if len(f) < 2 or not any(f): continue
            esperado = solo("".join(p for c in f for p in c))
            if esperado and esperado not in filas_md and not any(esperado in x for x in filas_md):
                filas.append("tabla %d, fila %d: %s" % (t, i, " | ".join(" ".join(c)[:25] for c in f)))
    return parrafos, filas

if __name__ == "__main__":
    md = convertir(sys.argv[1], sys.argv[2])
    print(sys.argv[2].split("/")[-1] + ": " + str(len(md)) + " chars, " + str(md.count(chr(10))+1) + " líneas")
    faltan = verificar(sys.argv[1], sys.argv[2])
    if faltan:
        print("  AVISO: pandoc descartó texto del .docx. Palabras sin aparecer en el .md:")
        print("  " + ", ".join(faltan[:30]))
    else:
        print("  verificado: no falta texto del .docx")
    parrafos, filas = verificar_estructura(sys.argv[1], sys.argv[2])
    if parrafos:
        print("  AVISO: párrafos del .docx sin reflejo en el .md: " + str(len(parrafos)))
        for p in parrafos[:15]: print("    " + p)
    if filas:
        print("  AVISO: filas de tabla con celdas perdidas o cambiadas: " + str(len(filas)))
        for f in filas[:15]: print("    " + f)
    if not parrafos and not filas:
        print("  verificado: párrafos y celdas de tabla presentes en el .md")
