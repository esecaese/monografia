---
# Datos de la carátula (Resolución N.º 1055/2024, Anexo, parte pre textual, punto 1).
# El script build/generar_docx.py lee ESTE bloque y arma la portada con la
# tipografía exacta exigida. Los campos marcados con `_24pt` se componen en
# Times New Roman 24; el resto, en Times New Roman 12.
#
# `mes_anio_24pt` vale `auto`: el build lo resuelve al mes y año en que se corre.
# Para congelar la fecha de la entrega, reemplazar `auto` por el texto literal
# (por ejemplo, "Noviembre de 2026").

institucion_24pt:
  - "Ministerio de Relaciones Exteriores"
  - "Academia Diplomática y Consular Carlos Antonio López"

curso: "Curso de Perfeccionamiento"

fundamento: "Artículo 110 de la Ley N.º 6935/2022"

titulo_24pt: "Propuesta de automatización de la liquidación de aranceles consulares mediante inteligencia artificial en el Ministerio de Relaciones Exteriores del Paraguay"

autor: "Sven Knutson Sachelaridi"

ciudad: "Asunción"

mes_anio_24pt: "auto"
---

<!--
Este archivo no produce texto corrido: su único contenido efectivo es el bloque
YAML de arriba. La portada se compone programáticamente para poder respetar los
tamaños de fuente diferenciados (24 pt / 12 pt) que exige la Resolución.

Datos del alumno, para referencia:
  - Categoría en el Escalafón: Segundo Oficial en Administración y Asuntos Técnicos
  - Dependencia laboral: Dirección de Informática
  - Correo electrónico: sknutson@mre.gov.py
-->
