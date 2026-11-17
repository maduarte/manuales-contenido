"""Reconoce las alternativas de opción múltiple en un guion .md.

Convención del proyecto (2026-10-15):
- En el .md las alternativas van en líneas (párrafos) separadas, SIN numeración ni letras.
- En el Doc / .docx van SIEMPRE con letra: a. b. c. d.

Cómo se reconocen sin marcas: una pregunta de opción múltiple termina en una línea de clave con
letras (`IPP: Correcta {b}`, `IPP: Correctas {a, c}`, `Respuesta correcta: {b}`). Las alternativas son
los párrafos que van entre el enunciado (la última línea que termina en "?" o ":") y esa clave.
Una alternativa puede ser un árbol del plugin: la línea IPP que lo nombra más la línea JSON que sigue.

Si el bloque es ambiguo (aparece un encabezado, una tabla o una marca IPP suelta antes del
enunciado, o hay menos alternativas que la letra más alta de la clave) no se marca nada y se avisa
por stderr, para revisar a mano.
"""
import re
import sys

CLAVE = re.compile(
    r"^\s*(?:IPP:\s*Correctas?|Respuesta\s+correcta:?)\s*\{\s*([a-eA-E](?:\s*(?:,|;|y)\s*[a-eA-E])*)\s*\}\s*$")
PREFIJO = re.compile(r"^(\d{1,2}|[a-eA-E])[\.\)]\s+")
LETRAS = "abcde"


def quitar_prefijo(linea):
    """'1.  texto' / 'a. texto' / 'b) texto' -> 'texto', conservando la sangría."""
    sangria = linea[: len(linea) - len(linea.lstrip())]
    return sangria + PREFIJO.sub("", linea.lstrip(), count=1)


def _plano(linea):
    return linea.strip().replace("*", "").strip()


def detectar(lineas, avisar=True):
    """Devuelve {índice de línea: letra} con la primera línea de cada alternativa.

    `lineas` es la lista de líneas crudas del .md (con las vacías).
    """
    marcas = {}
    for k, linea in enumerate(lineas):
        m = CLAVE.match(linea)
        if not m:
            continue
        claves = [c.lower() for c in re.findall(r"[a-eA-E]", m.group(1))]
        unidades = []          # índices de la primera línea de cada alternativa, de abajo hacia arriba
        j = k - 1
        ok = False
        while j >= 0:
            s = lineas[j].strip()
            if not s:
                j -= 1
                continue
            if s.startswith('{"mode"'):              # árbol: va con la línea IPP que lo nombra
                h = j - 1
                while h >= 0 and not lineas[h].strip():
                    h -= 1
                if h < 0 or not quitar_prefijo(lineas[h].strip()).startswith("IPP"):
                    break
                unidades.append(h)
                j = h - 1
                continue
            p = _plano(quitar_prefijo(s))
            if p.endswith("?") or p.endswith(":"):
                ok = True
                break
            if (s.startswith("|") or s.startswith("#") or s == "---" or "EJERCICIO" in s
                    or quitar_prefijo(s).startswith(("IPP", "IPD", "IPM", "IPC"))):
                break
            unidades.append(j)
            j -= 1
        idx_max = max(LETRAS.index(c) for c in claves)
        if not ok or len(unidades) < 2 or len(unidades) <= idx_max:
            if avisar:
                print(f"[alternativas] línea {k + 1}: no se reconocieron las alternativas de "
                      f"'{linea.strip()}'; revisar a mano.", file=sys.stderr)
            continue
        if len(unidades) > len(LETRAS):
            if avisar:
                print(f"[alternativas] línea {k + 1}: {len(unidades)} alternativas, más que letras.",
                      file=sys.stderr)
            continue
        for n, idx in enumerate(reversed(unidades)):
            marcas[idx] = LETRAS[n]
    return marcas


def limpiar_md(texto):
    """Quita numeración y letras de las alternativas reconocidas (formato .md del proyecto)."""
    lineas = texto.split("\n")
    for idx in detectar(lineas):
        lineas[idx] = quitar_prefijo(lineas[idx])
    return "\n".join(lineas)


if __name__ == "__main__":
    # python3 alternativas.py guion.md   -> reescribe el .md sin numeración en las alternativas
    from pathlib import Path
    for ruta in sys.argv[1:]:
        p = Path(ruta)
        antes = p.read_text(encoding="utf-8")
        despues = limpiar_md(antes)
        if despues != antes:
            p.write_text(despues, encoding="utf-8")
        n = len(detectar(despues.split("\n"), avisar=False))
        print(f"{p.name}: {n} alternativas reconocidas")
