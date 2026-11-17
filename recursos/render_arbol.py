#!/usr/bin/env python3
"""Renderiza árboles del plugin de Moodle (JSON) como PNG, para poder ver el código.

Uso:
    python3 02_Herramientas/render_arbol.py arbol.json                 # un JSON suelto
    python3 02_Herramientas/render_arbol.py U3T2L9_guion.md            # todos los DIAG_Pnn de un guion
    python3 02_Herramientas/render_arbol.py U3T2L9_guion.md -o salida/ # carpeta de salida
    python3 02_Herramientas/render_arbol.py U3T2L9_guion.md --solo DIAG_P14,DIAG_P18
    python3 02_Herramientas/render_arbol.py U3T2L9_guion.md --hoja     # además, una hoja de contacto

Del guion se extrae cada bloque "<nombre DIAG_...>" seguido de una línea JSON (con o sin
"Árbol que se muestra" delante). Sale un PNG por árbol, con el nombre del DIAG.

Qué dibuja: nodos (círculo o redondeado, con su color), aristas con su etiqueta (probabilidad del
tramo), etiquetas flotantes (probabilidad de la rama), rótulos de etapa a la izquierda
(levelLabelOverrides) y la grilla. Los campos por completar se muestran así:
    {{T_INPUT}}      -> caja punteada con "?"
    [[a;b;c]]        -> "▾ a | b | c" (selector)
En modo árbol se asume que (left, top) es el centro del nodo; el plugin recalcula la disposición, así que el
orden relativo es el del plugin, pero la posición exacta de pantalla puede variar unos píxeles. En modo grafo
("mode": "graph") el plugin respeta las posiciones: (left, top) es la esquina superior izquierda, las aristas
van de centro a centro y no hay rótulos de etapa (dibujar_grafo). En modo árbol con orientación horizontal, la
disposición se calcula desde las aristas, como el plugin (dibujar_horizontal).
"""
import argparse
import json
import os
import re
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Circle, Rectangle
from matplotlib import font_manager

MORADO = "#423475"
GRIS = "#8b8a99"
GRILLA = "#eceef2"
FONDO = "#f9fafb"

for _f in ("Geist", "Inter", "Helvetica Neue", "DejaVu Sans"):
    if any(_f == x.name for x in font_manager.fontManager.ttflist):
        plt.rcParams["font.family"] = _f
        break


def es_oscuro(hexcolor):
    h = (hexcolor or MORADO).lstrip("#")
    if len(h) != 6:
        return True
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    return (0.299 * r + 0.587 * g + 0.114 * b) < 150


def texto_visible(t):
    """Traduce los tokens del plugin a algo legible. Devuelve (texto, es_campo)."""
    if t is None:
        return "", False
    if "{{T_INPUT}}" in t:
        return "?", True
    m = re.fullmatch(r"\[\[(.*)\]\]", t.strip())
    if m:
        return "▾ " + " | ".join(m.group(1).split(";")), True
    return t, False


def extraer_arboles(ruta):
    """Devuelve [(nombre, dict)] desde un .json suelto o un guion .md."""
    txt = open(ruta, encoding="utf-8").read()
    if ruta.lower().endswith(".json"):
        return [(os.path.splitext(os.path.basename(ruta))[0], json.loads(txt))]
    out = []
    for m in re.finditer(r"(DIAG_\w+)[ \t]*\n[ \t]*(\{\"mode\".*\})[ \t]*\n", txt):
        out.append((m.group(1), json.loads(m.group(2))))
    if not out:
        for m in re.finditer(r"(DIAG_\w+)\s*(\{\"mode\".*\})", txt):
            out.append((m.group(1), json.loads(m.group(2))))
    return out


def tamano_grafo(n):
    """Ancho y alto (px) de un nodo como los dibuja el plugin (styles.css de moodle-plugin-graphsandtrees)."""
    if n.get("shape") == "circle":
        return 67, 67
    t, campo = texto_visible(n.get("text", ""))
    if "[[" in n.get("text", ""):
        return 142, 55
    if campo:
        return 160, 55
    return max(84, 6.6 * len(t)) + 39, 51


def dibujar_grafo(d, nombre="", destino=None, escala=1.0, ax=None, titulo=True):
    """Modo grafo: el plugin respeta las posiciones del JSON, (left, top) es la esquina superior izquierda del
    nodo, las aristas rectas van de centro a centro y no se dibujan los rótulos de etapa."""
    caja = {n["id"]: (n["left"], n["top"]) + tamano_grafo(n) for n in d["nodes"]}
    x0 = min(l for l, t, w, h in caja.values()) - 30
    y0 = min(t for l, t, w, h in caja.values()) - 30
    x1 = max(l + w for l, t, w, h in caja.values()) + 30
    y1 = max(t + h for l, t, w, h in caja.values()) + 30
    W, H = x1 - x0, y1 - y0
    propia = ax is None
    if propia:
        fig, ax = plt.subplots(figsize=(W / 100 * escala, H / 100 * escala), dpi=140)
        fig.patch.set_facecolor(FONDO)
    ax.set_xlim(x0, x1)
    ax.set_ylim(y1, y0)
    ax.set_aspect("equal")
    ax.axis("off")
    ax.add_patch(Rectangle((x0, y0), W, H, color=FONDO, zorder=0))
    if d.get("showGrid", True):
        g = int(x0 // 40) * 40
        while g <= x1:
            ax.plot([g, g], [y0, y1], color=GRILLA, lw=0.6, zorder=0.5)
            g += 40
        g = int(y0 // 40) * 40
        while g <= y1:
            ax.plot([x0, x1], [g, g], color=GRILLA, lw=0.6, zorder=0.5)
            g += 40
    for e in d["edges"]:
        (la, ta, wa, ha), (lb, tb, wb, hb) = caja[e["from"]], caja[e["to"]]
        xa, ya, xb, yb = la + wa / 2, ta + ha / 2, lb + wb / 2, tb + hb / 2
        ax.plot([xa, xb], [ya, yb], color=GRIS, lw=1.1, zorder=1)
        if e.get("label", ""):
            t, campo = texto_visible(e["label"])
            ax.text((xa + xb) / 2, (ya + yb) / 2, t, fontsize=8.5, ha="center", va="center", zorder=3,
                    color=e.get("textcolor") or "#222",
                    bbox=dict(boxstyle="round,pad=0.25", fc=("#fff3cd" if campo else FONDO),
                              ec=("#c9a227" if campo else "none"), ls="--" if campo else "-", lw=0.8))
    for n in d["nodes"]:
        l, t_, w, h = caja[n["id"]]
        col = n.get("color") or MORADO
        borde = n.get("bordercolor") or col
        if n.get("shape") == "circle":
            ax.add_patch(Circle((l + w / 2, t_ + h / 2), w / 2, fc=col, ec=borde, lw=1.2, zorder=2))
        else:
            r = 0 if n.get("shape") == "rectangle" else 8
            ax.add_patch(FancyBboxPatch((l, t_), w, h, boxstyle="round,pad=0,rounding_size=%d" % r,
                                        fc=col, ec=borde, lw=1.2, zorder=2))
        txt, campo = texto_visible(n.get("text", ""))
        colt = n.get("textcolor") or ("#ffffff" if es_oscuro(col) else "#111111")
        if txt.strip():
            ax.text(l + w / 2, t_ + h / 2, txt, fontsize=7.4, color=colt, ha="center", va="center", zorder=4,
                    bbox=(dict(boxstyle="round,pad=0.3", fc="#fff3cd", ec="#c9a227", ls="--", lw=0.8) if campo else None))
    if titulo and nombre:
        ax.text(x1 - 6, y0 + 12, nombre, fontsize=7.5, color=GRIS, ha="right", va="center")
    if propia:
        fig.tight_layout(pad=0.1)
        if destino:
            fig.savefig(destino, facecolor=fig.get_facecolor())
            plt.close(fig)
        return fig
    return None


def dibujar_horizontal(d, nombre="", destino=None, escala=1.0, ax=None, titulo=True):
    """Modo árbol con "orientation": "horizontal". Como el plugin (relayoutTree), la disposición sale de las
    aristas y no de (left, top): la raíz a la izquierda, un nivel por columna, las hojas apiladas de arriba hacia
    abajo y cada padre a la altura media de sus hijos. La etiqueta flotante va a la derecha de la hoja y los rótulos
    de etapa (levelLabelOverrides) arriba de cada columna."""
    nodos = {n["id"]: n for n in d["nodes"]}
    hijos = {}
    for e in d["edges"]:
        hijos.setdefault(e["from"], []).append(e["to"])
    con_padre = {e["to"] for e in d["edges"]}
    raices = [i for i in nodos if i not in con_padre and (i in hijos or len(nodos) == 1)]
    prof, fila, orden = {}, {}, []

    def recorrer(i, p):
        prof[i] = p
        hs = hijos.get(i, [])
        if not hs:
            fila[i] = len(orden)
            orden.append(i)
        else:
            for h in hs:
                recorrer(h, p + 1)
            fila[i] = sum(fila[h] for h in hs) / len(hs)

    for r in raices:
        recorrer(r, 0)

    def ancho(n):
        if n.get("shape") == "circle":
            return 44
        t, _ = texto_visible(n.get("text", ""))
        return max(90, 7.2 * len(t) + 28)

    alto, paso_y = 34, 46
    maxp = max(prof.values())
    anchos = [max((ancho(nodos[i]) for i in prof if prof[i] == k), default=60) for k in range(maxp + 1)]
    x_col, x = [], 30
    for k in range(maxp + 1):
        x_col.append(x)
        x += anchos[k] + 90
    x_fin = x + 40
    y0 = 70
    pos = {i: (x_col[prof[i]] + anchos[prof[i]] / 2, y0 + fila[i] * paso_y) for i in prof}
    y_fin = y0 + (len(orden) - 1) * paso_y + 50
    W, H = x_fin, y_fin
    propia = ax is None
    if propia:
        fig, ax = plt.subplots(figsize=(W / 100 * escala, H / 100 * escala), dpi=140)
        fig.patch.set_facecolor(FONDO)
    ax.set_xlim(0, W)
    ax.set_ylim(H, 0)
    ax.set_aspect("equal")
    ax.axis("off")
    ax.add_patch(Rectangle((0, 0), W, H, color=FONDO, zorder=0))
    if d.get("showLevelLabels"):
        ov = d.get("levelLabelOverrides", [])
        for k in range(maxp + 1):
            o = ov[k] if k < len(ov) else {}
            if o.get("show") is False:
                continue
            t = ": ".join(x for x in (o.get("title", "Etapa %d" % k if k else ""), o.get("subtitle", "")) if x)
            if t:
                ax.text(x_col[k] + anchos[k] / 2, 28, t, fontsize=8, color=GRIS, weight="bold", ha="center")
    for e in d["edges"]:
        (xa, ya), (xb, yb) = pos[e["from"]], pos[e["to"]]
        xa += (22 if nodos[e["from"]].get("shape") == "circle" else ancho(nodos[e["from"]]) / 2)
        xb -= ancho(nodos[e["to"]]) / 2
        ax.plot([xa, xb], [ya, yb], color=GRIS, lw=1.0, zorder=1)
        if e.get("label", ""):
            t, campo = texto_visible(e["label"])
            ax.text((xa + xb) / 2, (ya + yb) / 2, t, fontsize=8, ha="center", va="center", zorder=3,
                    bbox=dict(boxstyle="round,pad=0.25", fc=("#fff3cd" if campo else FONDO),
                              ec=("#c9a227" if campo else "none"), ls="--" if campo else "-", lw=0.8))
    for i, (x, y) in pos.items():
        n = nodos[i]
        col = n.get("color") or MORADO
        borde = n.get("bordercolor") or col
        t, campo = texto_visible(n.get("text", ""))
        if n.get("shape") == "circle":
            ax.add_patch(Circle((x, y), 18, fc=col, ec=borde, lw=1.2, zorder=2))
        else:
            w = ancho(n)
            ax.add_patch(FancyBboxPatch((x - w / 2, y - alto / 2), w, alto, boxstyle="round,pad=0,rounding_size=7",
                                        fc=col, ec=borde, lw=1.2, zorder=2))
        if t.strip():
            ax.text(x, y, t, fontsize=7.6, color=n.get("textcolor") or ("#ffffff" if es_oscuro(col) else "#111111"),
                    ha="center", va="center", zorder=4,
                    bbox=(dict(boxstyle="round,pad=0.3", fc="#fff3cd", ec="#c9a227", ls="--", lw=0.8) if campo else None))
        fl = n.get("floatingLabel", "")
        if fl and i not in hijos:
            ft, fcampo = texto_visible(fl)
            ax.text(x + ancho(n) / 2 + 10, y, ft, fontsize=8.5, ha="left", va="center", zorder=4, weight="bold",
                    color=n.get("floatingLabelTextColor") or "#222",
                    bbox=dict(boxstyle="round,pad=0.3", fc=("#fff3cd" if fcampo else "#ffffff"),
                              ec=("#c9a227" if fcampo else GRIS), ls="--" if fcampo else "-", lw=0.8))
    if titulo and nombre:
        ax.text(W - 6, 12, nombre, fontsize=7.5, color=GRIS, ha="right", va="center")
    if propia:
        fig.tight_layout(pad=0.1)
        if destino:
            fig.savefig(destino, facecolor=fig.get_facecolor())
            plt.close(fig)
        return fig
    return None


def dibujar(d, nombre="", destino=None, escala=1.0, ax=None, titulo=True):
    if d.get("mode") == "graph":
        return dibujar_grafo(d, nombre, destino, escala, ax, titulo)
    if d.get("orientation") == "horizontal":
        return dibujar_horizontal(d, nombre, destino, escala, ax, titulo)
    nodos = {n["id"]: n for n in d["nodes"]}
    aristas = d["edges"]
    lefts = [n["left"] for n in nodos.values()]
    tops = sorted({n["top"] for n in nodos.values()})

    def ancho(n):
        if n.get("shape") == "circle":
            return 44
        t, _ = texto_visible(n.get("text", ""))
        return max(70, 8.0 * len(t) + 30)

    alto = 38
    margen_izq = 150 if d.get("showLevelLabels") else 40
    x_min = min(n["left"] - ancho(n) / 2 for n in nodos.values()) - 40
    x_max = max(n["left"] + ancho(n) / 2 for n in nodos.values()) + 40
    y_min = min(tops) - 50
    y_max = max(tops) + 90
    x0 = x_min - (margen_izq if d.get("showLevelLabels") else 0)
    W, H = x_max - x0, y_max - y_min

    propia = ax is None
    if propia:
        fig, ax = plt.subplots(figsize=(W / 100 * escala, H / 100 * escala), dpi=140)
        fig.patch.set_facecolor(FONDO)
    ax.set_xlim(x0, x_max)
    ax.set_ylim(y_max, y_min)  # y hacia abajo
    ax.set_aspect("equal")
    ax.axis("off")
    ax.add_patch(Rectangle((x0, y_min), W, H, color=FONDO, zorder=0))

    if d.get("showGrid", True):
        paso = 40
        gx = int(x0 // paso) * paso
        while gx <= x_max:
            ax.plot([gx, gx], [y_min, y_max], color=GRILLA, lw=0.6, zorder=0.5)
            gx += paso
        gy = int(y_min // paso) * paso
        while gy <= y_max:
            ax.plot([x0, x_max], [gy, gy], color=GRILLA, lw=0.6, zorder=0.5)
            gy += paso

    # rótulos de etapa
    if d.get("showLevelLabels"):
        ov = d.get("levelLabelOverrides", [])
        tam = int(d.get("levelLabelSize", 24)) * 0.55
        for i, y in enumerate(tops):
            o = ov[i] if i < len(ov) else {}
            if o.get("show") is False:
                continue
            titulo_e = o.get("title", "Etapa %d" % i if i else "")
            sub = o.get("subtitle", "")
            xr = x0 + 18
            if titulo_e:
                ax.text(xr, y - 8, titulo_e, fontsize=tam * 0.62, color=GRIS, weight="bold", va="center", ha="left")
            if sub:
                ax.text(xr, y + 12, sub, fontsize=tam * 0.5, color=GRIS, va="center", ha="left")

    # aristas
    for e in aristas:
        a, b = nodos[e["from"]], nodos[e["to"]]
        xa, ya = a["left"], a["top"] + (22 if a.get("shape") == "circle" else alto / 2)
        xb, yb = b["left"], b["top"] - alto / 2
        ax.plot([xa, xb], [ya, yb], color=GRIS, lw=1.1, zorder=1)
        et = e.get("label", "")
        if et:
            t, campo = texto_visible(et)
            xm, ym = (xa + xb) / 2, (ya + yb) / 2
            ax.text(xm, ym, t, fontsize=8.5, ha="center", va="center", color=e.get("textcolor") or "#222",
                    zorder=3, bbox=dict(boxstyle="round,pad=0.25", fc=("#fff3cd" if campo else FONDO),
                                        ec=("#c9a227" if campo else "none"), ls="--" if campo else "-", lw=0.8))

    # nodos
    for n in nodos.values():
        x, y = n["left"], n["top"]
        col = n.get("color") or MORADO
        borde = n.get("bordercolor") or col
        t, campo = texto_visible(n.get("text", ""))
        if n.get("shape") == "circle":
            ax.add_patch(Circle((x, y), 22, fc=col, ec=borde, lw=1.2, zorder=2))
        else:
            w = ancho(n)
            ax.add_patch(FancyBboxPatch((x - w / 2, y - alto / 2), w, alto,
                                        boxstyle="round,pad=0,rounding_size=7",
                                        fc=col, ec=borde, lw=1.2, zorder=2))
        colt = n.get("textcolor") or ("#ffffff" if es_oscuro(col) else "#111111")
        if t.strip():
            ax.text(x, y, t, fontsize=8.5, color=colt, ha="center", va="center", zorder=4)
        fl = n.get("floatingLabel", "")
        if fl:
            ft, fcampo = texto_visible(fl)
            ax.text(x, y + alto / 2 + 16, ft, fontsize=9, ha="center", va="center", zorder=4,
                    color=n.get("floatingLabelTextColor") or "#222", weight="bold",
                    bbox=dict(boxstyle="round,pad=0.3", fc=("#fff3cd" if fcampo else "#ffffff"),
                              ec=("#c9a227" if fcampo else GRIS), ls="--" if fcampo else "-", lw=0.8))
    if titulo and nombre:
        ax.text(x_max - 6, y_min + 12, nombre, fontsize=7.5, color=GRIS, ha="right", va="center")

    if propia:
        fig.tight_layout(pad=0.1)
        if destino:
            fig.savefig(destino, facecolor=fig.get_facecolor())
            plt.close(fig)
        return fig
    return None


def main():
    ap = argparse.ArgumentParser(description="Renderiza árboles del plugin (JSON) a PNG.")
    ap.add_argument("entrada", help=".json suelto o guion .md con bloques DIAG_Pnn")
    ap.add_argument("-o", "--salida", help="carpeta de salida (por defecto, junto al archivo, en 'arboles_png/')")
    ap.add_argument("--solo", help="nombres DIAG separados por coma")
    ap.add_argument("--escala", type=float, default=1.0)
    ap.add_argument("--hoja", action="store_true", help="además genera hoja_de_contacto.png")
    a = ap.parse_args()

    arboles = extraer_arboles(a.entrada)
    if a.solo:
        pedidos = {s.strip() for s in a.solo.split(",")}
        arboles = [(n, d) for n, d in arboles if n in pedidos]
    if not arboles:
        sys.exit("No se encontró ningún árbol.")
    carpeta = a.salida or os.path.join(os.path.dirname(os.path.abspath(a.entrada)), "arboles_png")
    os.makedirs(carpeta, exist_ok=True)
    rutas = []
    for nombre, d in arboles:
        ruta = os.path.join(carpeta, nombre + ".png")
        dibujar(d, nombre, ruta, a.escala)
        rutas.append(ruta)
        print(ruta)
    if a.hoja:
        from PIL import Image
        ims = [Image.open(r).convert("RGB") for r in rutas]
        ancho = max(i.width for i in ims)
        hoja = Image.new("RGB", (ancho, sum(i.height + 10 for i in ims)), "white")
        y = 0
        for i in ims:
            hoja.paste(i, (0, y))
            y += i.height + 10
        ruta = os.path.join(carpeta, "hoja_de_contacto.png")
        hoja.save(ruta)
        print(ruta)


if __name__ == "__main__":
    main()
