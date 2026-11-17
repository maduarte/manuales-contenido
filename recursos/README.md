# Recursos: conversores de guiones

Scripts en Python para trabajar con los guiones de lección del proyecto.

| Archivo | Para qué |
|---|---|
| `leccion_a_html.py` | Convierte un guion (`.docx` bajado del Google Doc, o `.md`) en una página HTML de lectura, con las preguntas que se pueden comprobar, los diálogos, las diapositivas y los árboles. |
| `doc_a_md.py` | Baja un Google Doc (exportado como `.docx`) a `.md` con el formato del proyecto. |
| `alternativas.py` | Módulo auxiliar: reconoce las alternativas de opción múltiple. Lo usan los dos conversores. |
| `encabezado_corto.py` | Módulo auxiliar: lee el encabezado corto de cada componente. Lo usa `leccion_a_html.py`. |
| `render_arbol.py` | Módulo auxiliar: dibuja los diagramas de árbol del plugin. Lo usa `leccion_a_html.py`. |

Los cinco archivos deben estar en la misma carpeta.

## Uso

```bash
# Guion a HTML
python3 leccion_a_html.py U3T3L1.docx -o U3T3L1_vista.html
python3 leccion_a_html.py U3T3L4_guion.md -o U3T3L4_vista.html

# Google Doc (.docx) a markdown
python3 doc_a_md.py U3T3L1.docx U3T3L1_guion.md
```

## Requisitos

- Python 3.9 o más reciente.
- Para `leccion_a_html.py`: `pip install python-docx pillow matplotlib`.
- Para `doc_a_md.py`: [pandoc](https://pandoc.org/installing.html) instalado en el sistema.
