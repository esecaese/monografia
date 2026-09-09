# -*- coding: utf-8 -*-
"""
Genera `referencia/plantilla_formato.docx`: el documento base de estilos.

Equivale al `reference-doc` que se usaría con Pandoc, pero construido con
python-docx para poder fijar con exactitud las medidas de la Resolución
N.º 1055/2024 en lugar de heredar los valores por defecto de Word.

Uso:
    python build/crear_plantilla.py

`generar_docx.py` lo invoca solo si la plantilla no existe todavía.
"""

from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

from formato import (
    ALINEACION_CUERPO,
    ALTO_HOJA,
    ANCHO_HOJA,
    ESPACIO_ENTRE_PARRAFOS,
    FUENTE,
    IDIOMA,
    INTERLINEADO,
    MARGEN_DERECHO,
    MARGEN_INFERIOR,
    MARGEN_IZQUIERDO,
    MARGEN_SUPERIOR,
    SANGRIA_BIBLIOGRAFIA,
    SANGRIA_CITA,
    TAM_CUERPO,
    TAM_NOTA_PIE,
    aplicar_fuente,
)

RAIZ = Path(__file__).resolve().parent.parent
DESTINO = RAIZ / "referencia" / "plantilla_formato.docx"

NEGRO = RGBColor(0x00, 0x00, 0x00)


def _parrafo(estilo, alineacion=None, interlineado=INTERLINEADO,
             espacio_despues=ESPACIO_ENTRE_PARRAFOS, espacio_antes=Pt(0)):
    pf = estilo.paragraph_format
    if alineacion is not None:
        pf.alignment = alineacion
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = interlineado
    pf.space_before = espacio_antes
    pf.space_after = espacio_despues
    return pf


def _fijar_idioma(estilo, idioma=IDIOMA):
    r_pr = estilo.element.get_or_add_rPr()
    lang = r_pr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        r_pr.append(lang)
    lang.set(qn("w:val"), idioma)


def _obtener_o_crear(documento, nombre, base=None):
    """Devuelve un estilo de párrafo, creándolo si Word lo tiene como latente."""
    try:
        return documento.styles[nombre]
    except KeyError:
        estilo = documento.styles.add_style(nombre, WD_STYLE_TYPE.PARAGRAPH)
        if base is not None:
            estilo.base_style = documento.styles[base]
        return estilo


def _configurar_pagina(documento):
    """A4, márgenes de la Resolución e impresión a una sola carilla."""
    for seccion in documento.sections:
        seccion.page_width = ANCHO_HOJA
        seccion.page_height = ALTO_HOJA
        seccion.top_margin = MARGEN_SUPERIOR
        seccion.bottom_margin = MARGEN_INFERIOR
        seccion.left_margin = MARGEN_IZQUIERDO
        seccion.right_margin = MARGEN_DERECHO
        seccion.gutter = Pt(0)
        # Una sola carilla por hoja: sin márgenes simétricos ni página impar/par.
        seccion.different_first_page_header_footer = False


def _estilo_normal(documento):
    estilo = documento.styles["Normal"]
    aplicar_fuente(estilo, FUENTE, TAM_CUERPO)
    estilo.font.color.rgb = NEGRO
    _fijar_idioma(estilo)
    _parrafo(estilo, ALINEACION_CUERPO)
    estilo.paragraph_format.widow_control = True


def _estilos_titulos(documento):
    """Heading 1-3: base de la tabla de contenido automática.

    Se conservan los estilos integrados (y su w:outlineLvl) para que el campo
    TOC los reconozca; solo se reemplaza su apariencia.
    """
    # Nivel 1: capítulos y títulos mayores. Mayúsculas, negrita, centrado.
    h1 = documento.styles["Heading 1"]
    aplicar_fuente(h1, FUENTE, TAM_CUERPO)
    h1.font.bold = True
    h1.font.italic = False
    h1.font.all_caps = False  # el texto ya se pasa en mayúsculas desde el build
    h1.font.color.rgb = NEGRO
    _fijar_idioma(h1)
    _parrafo(h1, WD_ALIGN_PARAGRAPH.CENTER,
             espacio_antes=Pt(0), espacio_despues=Pt(24))
    h1.paragraph_format.keep_with_next = True
    h1.paragraph_format.page_break_before = False  # lo decide el build

    # Niveles 2 y 3: subtítulos. Negrita, alineados a la izquierda,
    # con la caja tipográfica tal como fue escrita en el Markdown.
    for nombre, espacio_antes in (("Heading 2", Pt(18)), ("Heading 3", Pt(12))):
        estilo = documento.styles[nombre]
        aplicar_fuente(estilo, FUENTE, TAM_CUERPO)
        estilo.font.bold = True
        estilo.font.italic = False
        estilo.font.all_caps = False
        estilo.font.color.rgb = NEGRO
        _fijar_idioma(estilo)
        _parrafo(estilo, WD_ALIGN_PARAGRAPH.LEFT,
                 espacio_antes=espacio_antes, espacio_despues=Pt(12))
        estilo.paragraph_format.keep_with_next = True


def _estilos_indice(documento):
    """Estilos TOC 1-3, que Word aplica a las entradas del índice generado."""
    niveles = (
        ("TOC 1", True, 0),
        ("TOC 2", True, 1),
        ("TOC 3", True, 2),
    )
    for nombre, negrita, nivel in niveles:
        estilo = _obtener_o_crear(documento, nombre, base="Normal")
        aplicar_fuente(estilo, FUENTE, TAM_CUERPO)
        estilo.font.bold = negrita
        estilo.font.color.rgb = NEGRO
        _fijar_idioma(estilo)
        pf = _parrafo(estilo, WD_ALIGN_PARAGRAPH.LEFT, espacio_despues=Pt(6))
        pf.left_indent = SANGRIA_BIBLIOGRAFIA * nivel


def _estilos_propios(documento):
    # Dedicatoria: "centrada en la página, justificada al margen derecho".
    dedicatoria = _obtener_o_crear(documento, "Dedicatoria", base="Normal")
    aplicar_fuente(dedicatoria, FUENTE, TAM_CUERPO)
    _fijar_idioma(dedicatoria)
    _parrafo(dedicatoria, WD_ALIGN_PARAGRAPH.RIGHT)

    # Agradecimiento(s): texto centrado.
    agradecimiento = _obtener_o_crear(documento, "Agradecimiento", base="Normal")
    aplicar_fuente(agradecimiento, FUENTE, TAM_CUERPO)
    _fijar_idioma(agradecimiento)
    _parrafo(agradecimiento, WD_ALIGN_PARAGRAPH.CENTER)

    # Referencias bibliográficas: sangría francesa (APA 7.ª ed., 2019).
    # Se alinean a la izquierda: justificar entradas con URLs largas abre
    # huecos entre palabras que APA desaconseja.
    biblio = _obtener_o_crear(documento, "Bibliografia APA", base="Normal")
    aplicar_fuente(biblio, FUENTE, TAM_CUERPO)
    _fijar_idioma(biblio)
    pf = _parrafo(biblio, WD_ALIGN_PARAGRAPH.LEFT)
    pf.left_indent = SANGRIA_BIBLIOGRAFIA
    pf.first_line_indent = -SANGRIA_BIBLIOGRAFIA

    # Citas textuales de más de 40 palabras (APA 7): bloque sangrado, sin comillas.
    cita = _obtener_o_crear(documento, "Cita en bloque", base="Normal")
    aplicar_fuente(cita, FUENTE, TAM_CUERPO)
    _fijar_idioma(cita)
    pf = _parrafo(cita, ALINEACION_CUERPO)
    pf.left_indent = SANGRIA_CITA

    # Viñetas.
    vineta = _obtener_o_crear(documento, "List Bullet", base="Normal")
    aplicar_fuente(vineta, FUENTE, TAM_CUERPO)
    _fijar_idioma(vineta)
    pf = _parrafo(vineta, ALINEACION_CUERPO, espacio_despues=Pt(12))
    pf.left_indent = SANGRIA_BIBLIOGRAFIA
    pf.first_line_indent = -SANGRIA_BIBLIOGRAFIA

    # Carátula.
    for nombre, tam in (("Caratula 24", 24), ("Caratula 12", TAM_CUERPO)):
        estilo = _obtener_o_crear(documento, nombre, base="Normal")
        aplicar_fuente(estilo, FUENTE, tam)
        estilo.font.color.rgb = NEGRO
        _fijar_idioma(estilo)
        _parrafo(estilo, WD_ALIGN_PARAGRAPH.CENTER,
                 interlineado=1.0, espacio_despues=Pt(0))

    # Encabezado: aloja el número de página, centrado en la parte superior.
    encabezado = _obtener_o_crear(documento, "Header", base="Normal")
    aplicar_fuente(encabezado, FUENTE, TAM_CUERPO)
    _fijar_idioma(encabezado)
    _parrafo(encabezado, WD_ALIGN_PARAGRAPH.CENTER,
             interlineado=1.0, espacio_despues=Pt(0))

    # Notas al pie: "mismo tipo de fuente e interlineado, tamaño 10".
    for nombre in ("Footnote Text", "Endnote Text"):
        estilo = _obtener_o_crear(documento, nombre, base="Normal")
        aplicar_fuente(estilo, FUENTE, TAM_NOTA_PIE)
        estilo.font.color.rgb = NEGRO
        _fijar_idioma(estilo)
        _parrafo(estilo, ALINEACION_CUERPO, espacio_despues=Pt(6))

    for nombre in ("Footnote Reference", "Endnote Reference"):
        try:
            estilo = documento.styles[nombre]
        except KeyError:
            estilo = documento.styles.add_style(nombre, WD_STYLE_TYPE.CHARACTER)
        aplicar_fuente(estilo, FUENTE, TAM_NOTA_PIE)
        estilo.font.superscript = True


def construir_plantilla(destino=DESTINO):
    documento = Document()

    # El documento base de python-docx trae un párrafo vacío: se descarta para
    # que la plantilla sea puramente un contenedor de estilos.
    for parrafo in list(documento.paragraphs):
        parrafo._element.getparent().remove(parrafo._element)

    _configurar_pagina(documento)
    _estilo_normal(documento)
    _estilos_titulos(documento)
    _estilos_indice(documento)
    _estilos_propios(documento)

    destino.parent.mkdir(parents=True, exist_ok=True)
    documento.save(destino)
    return destino


if __name__ == "__main__":
    ruta = construir_plantilla()
    print(f"Plantilla de estilos generada: {ruta.relative_to(RAIZ)}")
