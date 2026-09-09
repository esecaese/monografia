# Monografía — Automatización de la liquidación de aranceles consulares mediante inteligencia artificial

Trabajo monográfico del **Curso de Perfeccionamiento** de la Academia Diplomática y Consular "Carlos Antonio López" (Ministerio de Relaciones Exteriores del Paraguay), presentado en el marco del artículo 110 de la Ley N.º 6935/22.

- **Alumno:** Sven Knutson Sachelaridi
- **Categoría en el Escalafón:** Segundo Oficial en Administración y Asuntos Técnicos
- **Dependencia:** Dirección de Informática
- **Correo:** sknutson@mre.gov.py

**Tema.** Propuesta de un asistente virtual integrado al portal de apostillas que evacúe consultas, guíe al usuario y calcule automáticamente el arancel de legalización/apostilla conforme a la Ley N.º 1.030/97 y la Ley N.º 4.987/13.

---

## 1. Cómo funciona este repositorio

El texto vive en archivos **Markdown**, uno por sección o capítulo. Esa es la **fuente de verdad**: es lo que se versiona, lo que se revisa y lo que se corrige.

El archivo Word se **genera** a partir de esos Markdown con un script que aplica las normas de formato de la **Resolución N.º 1055/2024** (Anexo). El `.docx` es un artefacto derivado y descartable; no se edita a mano salvo para la revisión final previa a la entrega.

```
monografia/
├── README.md
├── .gitignore
├── build/
│   ├── formato.py            # todas las constantes de la Resolución, en un solo lugar
│   ├── crear_plantilla.py    # genera referencia/plantilla_formato.docx
│   ├── generar_docx.py       # script principal de build
│   └── requirements.txt
├── referencia/
│   └── plantilla_formato.docx  # documento base de estilos (generado)
├── preliminares/
│   ├── caratula.md           # bloque YAML con los datos de portada
│   ├── dedicatoria.md
│   ├── agradecimiento.md
│   └── indice.md             # placeholder: el índice real es un campo TOC de Word
├── cuerpo/
│   ├── introduccion.md
│   ├── objetivos.md
│   ├── justificacion.md
│   ├── metodologia.md
│   ├── limitaciones.md
│   ├── capitulo-1.md … capitulo-5.md
│   └── conclusiones.md
├── postexto/
│   └── bibliografia.md
└── output/
    └── monografia-aranceles-consulares.docx   (generado, ignorado por git)
```

El orden en que se ensambla el documento está declarado en las listas `CARATULA`, `PRELIMINARES` y `CUERPO` al inicio de [`build/generar_docx.py`](build/generar_docx.py). Para agregar, quitar o reordenar una sección, se modifica **solo esa lista**.

---

## 2. Instalación de dependencias

La única dependencia es `python-docx`. **No hace falta Pandoc** (ver sección 4).

Ya están instaladas en este equipo (Python 3.12.10 y python-docx 1.2.0). Para reproducir el entorno en otra máquina:

```powershell
winget install --id Python.Python.3.12 -e
```

Cerrar y reabrir la terminal para que `python` quede en el PATH, y luego:

```powershell
python -m pip install -r build/requirements.txt
```

---

## 3. Generar el documento

Desde la raíz de `segundo-semestre/monografia`:

```powershell
python build/generar_docx.py
```

El script:

1. Genera `referencia/plantilla_formato.docx` si todavía no existe.
2. Concatena los Markdown en el orden del manifiesto.
3. Escribe `output/monografia-aranceles-consulares.docx`.
4. Imprime un informe de extensión por archivo, con la estimación de páginas del cuerpo computable y el equilibrio entre los capítulos redactados.

**Al abrir el .docx en Word:** aceptar la actualización de campos para que se construya el índice. Si no aparece el aviso, `Ctrl+E` (seleccionar todo) y luego `F9`.

Para regenerar la plantilla de estilos desde cero (por ejemplo, después de tocar `build/formato.py`):

```powershell
python build/crear_plantilla.py
```

---

## 4. Cumplimiento de la Resolución N.º 1055/2024

Todas las medidas están centralizadas en [`build/formato.py`](build/formato.py); cambiar una regla es cambiar una línea.

| Regla de la Resolución | Dónde se aplica |
|---|---|
| Hoja A4, una carilla | `crear_plantilla.py` → `_configurar_pagina` |
| Márgenes 2,5 cm sup./inf./der. y 3,5 cm izq. | `formato.py` → `MARGEN_*` |
| Times New Roman 12, interlineado 1,5 | estilo `Normal` |
| Espacio entre párrafos 1,5 | `ESPACIO_ENTRE_PARRAFOS` (18 pt = 1,5 × 12 pt) |
| Notas al pie: misma fuente, tamaño 10, interlineado 1,5 | estilos `Footnote Text` / `Endnote Text` |
| Carátula con bloques en 24 pt | `generar_docx.py` → `agregar_caratula`, leyendo el YAML de `caratula.md` |
| Parte pre textual en números romanos, centrados arriba | sección 1-2, `numerar_paginas(..., "lowerRoman")` |
| Cuerpo en números arábigos desde la Introducción, centrados arriba | sección 3, `numerar_paginas(..., "decimal")` |
| Encabezados de nivel 1 centrados, en mayúscula, negrita, sin punto final | `agregar_titulo` (nivel 1) + estilo `Heading 1` |
| Subtítulos en negrita, caja tal como se escribieron | estilos `Heading 2` / `Heading 3` |
| Índice como tabla de contenido automática | campo `TOC \o "1-3" \h \z \u` + `<w:updateFields/>` |
| El índice no se lista a sí mismo | el encabezado ÍNDICE usa el estilo `Titulo preliminar`, idéntico a la vista pero fuera del TOC |
| Dedicatoria justificada al margen derecho | estilo `Dedicatoria` |
| Agradecimiento centrado | estilo `Agradecimiento` |
| Bibliografía en orden alfabético, sangría francesa APA 7 | estilo `Bibliografia APA` |
| Extensión 20-50 páginas, excluida la parte pre textual y la bibliografía | informe del build (columna `*` = no computa) |
| Negritas solo en títulos, subtítulos y viñetas | los `.md` del cuerpo no usan `**` en texto corrido |

### Decisiones tomadas donde la Resolución no se pronuncia

Están todas marcadas con comentario en `build/formato.py` y se revierten en una línea:

- **Alineación del texto corrido:** justificada (`ALINEACION_CUERPO`).
- **Romanos en minúscula** (i, ii, iii) para la parte pre textual (`FORMATO_NUM_PRELIMINARES`); cambiar a `"upperRoman"` para I, II, III.
- **Carátula en la página i, sin número impreso;** la dedicatoria arranca en ii.
- **Bibliografía alineada a la izquierda** en lugar de justificada: las entradas con URLs largas abren huecos entre palabras que APA desaconseja.
- **La dedicatoria y los agradecimientos sí aparecen en el índice**, con su numeración romana; el índice no se lista a sí mismo. Para excluirlos también, aplicarles el estilo `Titulo preliminar` en `generar_docx.py`, igual que se hace con el encabezado ÍNDICE.

### Verificación realizada

El build se probó de punta a punta abriendo el resultado en Word: 52 comprobaciones estructurales sobre el `.docx` (secciones, `pgNumType`, márgenes, estilos, campos, orden y caja de los títulos) más la actualización real del campo TOC. El índice se genera con la paginación correcta —romanos en los preliminares (ii, iii), arábigos desde la Introducción (1)— y la carátula entra en una sola página.

### Por qué python-docx y no Pandoc

La Resolución exige tres cosas que Pandoc no puede expresar desde Markdown y que igual habría que post-procesar con `python-docx`:

1. Numeración romana en los preliminares y arábiga desde la Introducción, lo que requiere saltos de sección con su propio `<w:pgNumType>`.
2. Una carátula donde conviven bloques de 24 pt y de 12 pt.
3. Un campo TOC nativo de Word.

Resolverlo todo en un solo paso elimina una dependencia pesada y deja una única definición del formato.

---

## 5. Estado de avance

| Sección | Estado |
|---|---|
| Carátula | Estructura lista — **falta definir mes y año de presentación** |
| Dedicatoria | Redactada |
| Agradecimientos | **Borrador** — reemplazar por texto propio, o quitar del manifiesto (es opcional) |
| Índice | Automático (campo TOC) |
| Introducción | Redactada |
| Objetivos | Redactados (1 general + 5 específicos) |
| Justificación | Redactada |
| Metodología | Redactada |
| Limitaciones | Redactadas |
| Capítulo I — Marco normativo y contexto institucional | Redactado |
| Capítulo II — Diagnóstico del procedimiento actual | **Pendiente** (objetivo específico 2) |
| Capítulo III — Fundamentos de la automatización | **Pendiente** (objetivo específico 3) |
| Capítulo IV — Diseño del asistente virtual | **Pendiente** (objetivo específico 4) |
| Capítulo V — Viabilidad de la propuesta | **Pendiente** (objetivo específico 5) |
| Conclusiones y recomendaciones | **Pendiente** (se escribe al final) |
| Bibliografía | Compilada — **falta completar la ficha de Alfonso (1995)** |

Cada capítulo del cuerpo responde a **un único objetivo específico**, decisión metodológica ya tomada.

### Pendientes abiertos

- **Mes y año** de presentación en `preliminares/caratula.md`.
- **Ficha APA completa de Alfonso (1995)**, citado en Metodología y todavía sin referencia localizada.
- **Pregunta operativa heredada del Capítulo I:** bajo qué criterio tarifario se cobra hoy la Apostilla en la práctica, dado que ni la Ley N.º 4.987/13 ni el Decreto N.º 520/13 fijan una tasa expresa. Se responde en el Capítulo II.
- **`## OBJETIVOS` es un título de nivel 2**, no de nivel 1: así fluye a continuación de la Introducción, que termina anunciándolos, en vez de abrir página nueva. En el índice aparece anidado bajo INTRODUCCIÓN. Si se prefiere como sección independiente, cambiar `##` por `#` en `cuerpo/objetivos.md`.
- El Capítulo V debe organizarse en subtítulos explícitos de **viabilidad técnica / operativa / normativa**, conforme a la observación del profesor.
- **`### GENERAL` y `### ESPECÍFICOS` están en mayúsculas** en `cuerpo/objetivos.md` y así aparecen en el índice. La Resolución pide que los subtítulos lleven "primera letra mayúscula y el resto minúscula". El texto se dejó tal como fue redactado; para cumplir la regla al pie de la letra habría que escribirlos `### General` y `### Específicos`.
- **Extensión actual: 16 páginas de cuerpo computable** según Word (de la Introducción a las Conclusiones), sobre un mínimo de 20. Buena parte son páginas casi vacías de los capítulos pendientes. El informe del build da una estimación por recuento de palabras, siempre algo por debajo del recuento real de Word.

---

## 6. Convención de trabajo

Un commit por sección o capítulo redactado o corregido, con mensaje descriptivo:

```
Redacción completa del Capítulo II: diagnóstico del procedimiento actual
Corrección de estilo en la Justificación
Completar ficha APA de Alfonso (1995) en la bibliografía
```

Así el historial refleja la evolución del trabajo y permite volver a cualquier versión anterior de una sección.

### Sobre las tablas, gráficos y cuadros

La Resolución exige insertarlos **como imagen**. El flujo previsto: guardar la imagen en `referencia/` o en una carpeta `imagenes/`, e insertarla en el `.docx` final durante la revisión previa a la entrega. El script todavía no procesa imágenes desde Markdown; se agrega cuando haga falta.

---

## 7. Repositorio

Este directorio vive dentro del repositorio [`esecaese/academia`](https://github.com/esecaese/academia), que ya está conectado a GitHub. Para publicar cambios:

```bash
git push
```

No hace falta `git remote add origin`: el remoto ya está configurado.
