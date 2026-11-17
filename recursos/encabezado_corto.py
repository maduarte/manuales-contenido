#!/usr/bin/env python3
"""Encabezado de componente en formato corto (docente, 2026-11-12).

El docente escribe cada componente con tres líneas entre las líneas de guiones bajos:

    “Lección 1.4: Sistematización”
    Tipo de componente moodle: contenido.
    Tiempo aproximado: 10 min

Las herramientas (vista_estudiante.py, leccion_a_html.py, guion_a_docx.py) reconocen los componentes por el formato
largo (`IPM: Nº componente moodle de la clase. Tiempo aproximado: N min` y `nombre del componente : “…”`).
`normalizar()` traduce el formato corto al largo al leer el guion, así que los dos formatos funcionan. El ordinal
del componente se cuenta en el orden en que aparecen.
"""
import re

NOMBRE = re.compile(r'^\s*[“"](Lección[^”"]+)[”"]\s*$')
TIPO = re.compile(r"^\s*Tipo de componente moodle\s*:", re.I)
TIEMPO = re.compile(r"^\s*Tiempo aproximado\s*:\s*(.+?)\s*$", re.I)
IPM = re.compile(r"^\s*IPM\s*:\s*\d+\s*º\s*componente", re.I)


def normalizar(texto):
    lineas = texto.split("\n")
    salida, k, i = [], 0, 0
    while i < len(lineas):
        l = lineas[i]
        if IPM.match(l):
            k += 1
        m = NOMBRE.match(l)
        if m:
            # el bloque corto: nombre, tipo y tiempo en las líneas que siguen (sin contar las vacías)
            j, tipo, tiempo = i + 1, None, None
            while j < len(lineas) and j <= i + 6 and not set(lineas[j].strip()) == {"_"}:
                if TIPO.match(lineas[j]) and tipo is None:
                    tipo = lineas[j].strip()
                elif TIEMPO.match(lineas[j]) and tiempo is None:
                    tiempo = TIEMPO.match(lineas[j]).group(1)
                elif lineas[j].strip():
                    break
                j += 1
            if tipo:
                k += 1
                salida += [f"IPM: {k}º componente moodle de la clase. Tiempo aproximado: {tiempo or 'xx min'}", "",
                           f"nombre del componente : “{m.group(1)}”", "", tipo, ""]
                i = j
                continue
        salida.append(l)
        i += 1
    return "\n".join(salida)
