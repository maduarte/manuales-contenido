# Recursos para el equipo de contenido

Material de apoyo para escribir guiones de lección. Se lee como sitio web (GitHub Pages) desde `index.html`.

| Archivo | Qué es | Cómo se mantiene |
|---|---|---|
| `manual_diseno.html` | Manual: 16 pasos, 4 revisiones, listas de chequeo, glosario (mismo estilo que los otros dos) | A mano |
| `proceso_diseno_guiones.html` | Proceso de creación de un guion | Se genera localmente desde el proceso del proyecto |
| `U1_ruta.html` | Ruta de la Unidad 1 (análisis de gráficos) | Se genera desde la ruta del proyecto |
| `U2_ruta.html` | Ruta de la Unidad 2 (media, dispersión y correlación) | Se genera desde la ruta del proyecto |
| `U3_ruta.html` | Ruta de la Unidad 3 (probabilidad) | Se genera desde la ruta del proyecto |
| `U3_Tema3_ruta.html` | Redirige a la sección del Tema 3 dentro de `U3_ruta.html` | Se mantiene para no romper enlaces antiguos |
| `U4_ruta.html` | Ruta de la Unidad 4 (variables aleatorias) | Se genera desde la ruta del proyecto |
| `U5_ruta.html` | Ruta de la Unidad 5 (inferencia estadística) | Se genera desde la ruta del proyecto |
| `LDI_ruta_curso_implementado.html` | Ruta del curso LDI implementado (Límites, Derivadas e Integrales) | Se genera desde la ruta del proyecto |
| `LDI_decisiones_estructurales.html` | Decisiones estructurales del curso LDI nuevo | Se genera desde el proyecto |
| `ldi/` | Página de LDI: ruta del curso implementado, ruta nueva por unidad, fichas y avance de cada lección | Lo genera `02_Herramientas/generar_sitio_ldi.py --publicar` |
| `lecciones/` | Lecciones en construcción: una vista HTML por lección, como la verá el estudiante | La publica el skill producir-guion en cada entrega |
| `revision/` | Índice de los artifacts de revisión de cada lección (tabla de propósitos y vista comentable) | Lo publica `02_Herramientas/publicar_indice_artifacts.py` |

Los cambios se hacen en local; aquí se publica solo el HTML.
