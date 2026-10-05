# Monografía — Automatización de la liquidación de aranceles consulares mediante inteligencia artificial

Trabajo monográfico del **Curso de Perfeccionamiento** de la Academia Diplomática y Consular "Carlos Antonio López" (Ministerio de Relaciones Exteriores del Paraguay), presentado en el marco del artículo 110 de la Ley N.º 6935/22.

- **Alumno:** Sven Knutson Sachelaridi
- **Categoría en el Escalafón:** Segundo Oficial en Administración y Asuntos Técnicos
- **Dependencia:** Dirección de Informática
- **Correo:** sknutson@mre.gov.py

**Tema.** Propuesta de un asistente virtual integrado al portal de apostillas que evacúe consultas, guíe al usuario y calcule automáticamente el arancel de legalización/apostilla conforme a la Ley N.º 1.030/97 y la Ley N.º 4.987/13.

**Pregunta de investigación.** ¿En qué medida es viable automatizar la liquidación de aranceles consulares mediante un asistente virtual integrado al sistema de apostilla del MRE, bajo el marco normativo vigente y los principios de la gestión digital pública?

> **¿Retomando el trabajo en otro equipo?** Ver la sección 10, [Arrancar una sesión en cualquier equipo](#10-arrancar-una-sesión-en-cualquier-equipo).

---

## 1. Cómo funciona este repositorio

El texto vive en archivos **Markdown**, uno por sección o capítulo. Esa es la **fuente de verdad**: es lo que se versiona, lo que se revisa y lo que se corrige.

El archivo Word se **genera** a partir de esos Markdown con un script que aplica las normas de formato de la **Resolución N.º 1055/2024** (Anexo). El `.docx` es un artefacto derivado y descartable; no se edita a mano salvo para la revisión final previa a la entrega.

```
Monografia/
├── README.md
├── .gitignore
├── build/
│   ├── formato.py               # todas las constantes de la Resolución, en un solo lugar
│   ├── crear_plantilla.py       # genera referencia/plantilla_formato.docx
│   ├── generar_docx.py          # script principal de build
│   ├── verificar.py             # 55 comprobaciones sobre el .docx generado
│   ├── verificar_en_word.ps1    # abre en Word, actualiza el índice, cuenta páginas
│   └── requirements.txt
├── referencia/
│   └── plantilla_formato.docx   # documento base de estilos (generado, versionado)
├── preliminares/
│   ├── caratula.md              # bloque YAML con los datos de portada
│   ├── dedicatoria.md
│   ├── agradecimiento.md
│   └── indice.md                # placeholder: el índice real es un campo TOC de Word
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
├── fuentes/                     # copia local de todo lo que el trabajo cita
│   ├── README.md                # índice comentado: qué aporta cada fuente
│   ├── normativa/               # leyes y decretos
│   ├── academicas/              # papers e informes
│   ├── institucionales/         # páginas y manuales del MRE
│   └── capturas/                # extractos que no sobreviven como PDF
└── output/                      # todo generado, ignorado por git
    ├── monografia-aranceles-consulares.docx
    ├── vista-previa.pdf
    └── _combinado.md            # el Markdown concatenado, para depurar
```

El orden en que se ensambla el documento está declarado en las listas `CARATULA`, `PRELIMINARES` y `CUERPO` al inicio de [`build/generar_docx.py`](build/generar_docx.py). Para agregar, quitar o reordenar una sección, se modifica **solo esa lista**.

### Qué entiende el parser de Markdown

Es deliberadamente mínimo — reconoce lo que este trabajo usa y nada más:

| En el `.md` | En el `.docx` |
|---|---|
| `# Título` | nivel 1: negrita y página nueva. Mayúsculas y centrado desde los capítulos; en modo oración y a la izquierda en las secciones listadas en `TITULOS_EN_MODO_ORACION` |
| `## Subtítulo` / `### Sub-sub` | negrita, alineado a la izquierda, caja tal como se escribió |
| Párrafos separados por línea en blanco | texto corrido justificado |
| `- ` al inicio de línea | viñeta (o entrada de bibliografía con sangría francesa, en `bibliografia.md`) |
| `> ` al inicio de línea | cita en bloque sangrada (APA 7, más de 40 palabras) |
| `*cursiva*` y `**negrita**` | cursiva y negrita |
| `<!-- comentario -->` | **se descarta** — sirve para notas de trabajo que no deben llegar al documento |

Las viñetas deben caber en **una sola línea** del `.md`: el parser no une líneas dentro de una viñeta. El guion bajo no se interpreta como cursiva a propósito, porque aparece dentro de las URLs de la bibliografía.

---

## 2. Entorno y dependencias

La única dependencia es `python-docx`. **No hace falta Pandoc** (ver el porqué en la sección 5).

```powershell
winget install --id Python.Python.3.12 -e
python -m pip install -r build/requirements.txt
```

### Estado de este equipo (a septiembre de 2025)

Anotado porque cuesta más redescubrirlo que leerlo:

- **Python 3.12.10** instalado en `%LOCALAPPDATA%\Programs\Python\Python312\python.exe`. **No quedó en el PATH** de las terminales ya abiertas: si `python` no responde, usar la ruta completa o abrir una terminal nueva.
- **python-docx 1.2.0** instalado.
- **Microsoft Word** instalado — lo usa `build/verificar_en_word.ps1` por COM.
- **No hay Pandoc, ni poppler, ni LibreOffice.** No hacen falta.
- La identidad de Git (`esecaese` / `knut.sach@gmail.com`) está configurada **local a este repositorio**, no a nivel global.

---

## 3. Generar el documento

Desde la raíz del repositorio:

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

## 4. Verificar el resultado

Dos comprobaciones complementarias. Conviene correr las dos después de cualquier cambio al build.

### Sobre el XML del `.docx` — no necesita Word

```powershell
python build/verificar.py
```

55 comprobaciones: secciones y `pgNumType`, medidas de hoja y márgenes, estilos (fuentes, tamaños, interlineado, alineaciones), presencia del campo TOC y de `updateFields`, orden y caja de los títulos, y contenido (que los comentarios HTML no se hayan filtrado, que las cursivas se hayan convertido, cantidad de entradas de bibliografía y de viñetas). Devuelve 0 si todo pasa.

Los conteos esperados están declarados como constantes al inicio del archivo (`ENTRADAS_BIBLIOGRAFIA`, `VINETAS_OBJETIVOS`, `CANTIDAD_CAPITULOS`). **Cuando el contenido cambie a propósito, actualizarlos**; que fallen es justamente la señal de que algo se agregó o se perdió sin querer.

### Dentro de Word — lo que el XML no puede decir

```powershell
powershell -ExecutionPolicy Bypass -File build/verificar_en_word.ps1
```

Abre el documento, actualiza los campos de verdad, e informa el recuento **real** de páginas, el índice tal como queda generado, y en qué página física arranca cada título. Exporta además `output/vista-previa.pdf`. Cierra sin guardar: el `.docx` no se toca.

Sirve sobre todo para medir la extensión contra el mínimo de 20 páginas, que el estimador por palabras del build sólo aproxima.

> Si el script falla con una `COMException`, es que quedó una instancia de Word colgada. `Get-Process WINWORD | Stop-Process` y reintentar.

---

## 5. Cumplimiento de la Resolución N.º 1055/2024

Todas las medidas están centralizadas en [`build/formato.py`](build/formato.py); cambiar una regla es cambiar una línea.

| Regla de la Resolución | Dónde se aplica |
|---|---|
| Hoja A4, una carilla | `crear_plantilla.py` → `_configurar_pagina` |
| Márgenes 2,5 cm sup./inf./der. y 3,5 cm izq. | `formato.py` → `MARGEN_*` |
| Times New Roman 12, interlineado 1,5 | estilo `Normal` |
| Espacio entre párrafos 1,5 | `ESPACIO_ENTRE_PARRAFOS` (18 pt = 1,5 × 12 pt) |
| Notas al pie: misma fuente, tamaño 10, interlineado 1,5 | estilos `Footnote Text` / `Endnote Text` |
| Carátula con bloques en 24 pt | `generar_docx.py` → `agregar_caratula`, leyendo el YAML de `caratula.md` |
| Fecha de presentación en la carátula | `mes_anio_24pt: "auto"` → `resolver_mes_anio`, con los nombres de `MESES` |
| Parte pre textual en números romanos, centrados arriba | secciones 1-2, `numerar_paginas(..., "lowerRoman")` |
| Cuerpo en números arábigos desde la Introducción, centrados arriba | sección 3, `numerar_paginas(..., "decimal")` |
| Encabezados de nivel 1 centrados, en mayúscula, negrita, sin punto final | `agregar_titulo` (nivel 1) + estilo `Heading 1`. **Con la excepción indicada por el profesor** para las secciones previas a los capítulos —ver abajo |
| Subtítulos en negrita, caja tal como se escribieron | estilos `Heading 2` / `Heading 3` |
| Índice como tabla de contenido automática | campo `TOC \o "1-3" \h \z \u` + `<w:updateFields/>` |
| El índice no se lista a sí mismo | el encabezado ÍNDICE usa el estilo `Titulo preliminar` |
| Dedicatoria justificada al margen derecho | estilo `Dedicatoria` |
| Agradecimiento centrado | estilo `Agradecimiento` |
| Bibliografía en orden alfabético, sangría francesa APA 7 | estilo `Bibliografia APA` |
| Extensión 20-50 páginas, excluida la parte pre textual y la bibliografía | informe del build (columna `*` = no computa) |
| Negritas solo en títulos, subtítulos y viñetas | los `.md` del cuerpo no usan `**` en texto corrido |

### La trampa del campo TOC

Vale dejarla anotada porque no es evidente y cuesta un rato descubrirla: **el switch `\o "1-3"` del campo TOC selecciona los párrafos por ESTILO, no por nivel de esquema.** Bajarle el `w:outlineLvl` a un párrafo con estilo `Heading 1` **no** lo saca del índice.

Por eso el encabezado ÍNDICE lleva el estilo `Titulo preliminar`: visualmente idéntico a un título de nivel 1, pero al no ser un estilo de título queda fuera de la tabla de contenido. Sin eso, el índice se listaba a sí mismo.

### Decisiones tomadas donde la Resolución no se pronuncia

Están todas marcadas con comentario en `build/formato.py` y se revierten en una línea:

- **Alineación del texto corrido:** justificada (`ALINEACION_CUERPO`).
- **Romanos en minúscula** (i, ii, iii) para la parte pre textual (`FORMATO_NUM_PRELIMINARES`); cambiar a `"upperRoman"` para I, II, III.
- **Carátula en la página i, sin número impreso;** la dedicatoria arranca en ii.
- **Títulos de nivel 1 en modo oración** de la Introducción a las Limitaciones (`TITULOS_EN_MODO_ORACION`, en `generar_docx.py`). Es una **indicación expresa del profesor**, no una lectura de la Resolución, que para el nivel 1 pide mayúsculas y centrado sin excepciones. Se sigue al profesor porque es quien evalúa. Quitar una ruta del conjunto devuelve esa sección al formato de la Resolución.
- **El salto de página se conserva** en esas secciones: el profesor habló de caja y alineación, no de la paginación.
- **Nombre del mes en la carátula:** `"Setiembre"`, no `"Septiembre"` (`MESES`, en `formato.py`). Ambas formas son correctas; se eligió la habitual en la normativa paraguaya. Van con inicial mayúscula porque la fecha es una línea suelta de portada y no texto corrido.
- **Bibliografía alineada a la izquierda** en lugar de justificada: las entradas con URLs largas abren huecos entre palabras que APA desaconseja.
- **La dedicatoria y los agradecimientos sí aparecen en el índice**, con su numeración romana. Para excluirlos, aplicarles el estilo `Titulo preliminar` en `generar_docx.py`, igual que al encabezado ÍNDICE.

### Por qué python-docx y no Pandoc

La Resolución exige tres cosas que Pandoc no puede expresar desde Markdown y que igual habría que post-procesar con `python-docx`:

1. Numeración romana en los preliminares y arábiga desde la Introducción, lo que requiere saltos de sección con su propio `<w:pgNumType>`.
2. Una carátula donde conviven bloques de 24 pt y de 12 pt.
3. Un campo TOC nativo de Word.

Resolverlo todo en un solo paso elimina una dependencia pesada y deja una única definición del formato.

---

## 6. Estado de avance

| Sección | Estado |
|---|---|
| Carátula | Completa — la fecha se resuelve sola al generar |
| Dedicatoria | Redactada |
| Agradecimientos | **Borrador** — reemplazar por texto propio, o quitar del manifiesto (es opcional) |
| Índice | Automático (campo TOC) |
| Introducción | Redactada |
| Objetivos | Redactados (1 general + 5 específicos) |
| Justificación | Redactada |
| Metodología | Redactada |
| Limitaciones | Redactadas |
| Capítulo I — Marco normativo y contexto institucional | Redactado |
| Capítulo II — Diagnóstico del procedimiento actual | Redactado |
| Capítulo III — Fundamentos de la automatización | **Pendiente** (objetivo específico 3) |
| Capítulo IV — Diseño del asistente virtual | **Pendiente** (objetivo específico 4) |
| Capítulo V — Viabilidad de la propuesta | **Pendiente** (objetivo específico 5) |
| Conclusiones y recomendaciones | **Pendiente** (se escribe al final) |
| Bibliografía | Compilada — **falta completar la ficha de Alfonso (1995)** |

Cada capítulo del cuerpo responde a **un único objetivo específico**, decisión metodológica ya tomada.

**Extensión actual: 29 páginas totales, 21 de cuerpo computable** según Word (de la Introducción a las Conclusiones), sobre un mínimo de 20. El mínimo ya se alcanza con los capítulos III a V y las Conclusiones todavía vacíos.

### Pendientes abiertos

- **Congelar la fecha de la carátula** al entregar: hoy vale `auto` y sigue al reloj, de modo que el mes cambia si el documento se regenera más adelante. Antes de imprimir la versión definitiva, reemplazar `auto` por el texto literal en `preliminares/caratula.md`.
- **Ficha APA completa de Alfonso (1995)**, citado en Metodología y todavía sin referencia localizada.
- **Día y mes exactos del Decreto N.º 2129/14** para la ficha bibliográfica: el escaneo de SILpy los trae ilegibles por OCR. Todo indica el 26 de agosto de 2014; falta confirmarlo contra la copia oficial.
- **Dos fuentes citadas que faltan en `fuentes/`**: el informe del BID (Roseth, Reyes y Santiso, 2018) y la Recomendación de la UNESCO (2021). Ambos sitios rechazan las descargas automatizadas; hay que guardarlos a mano desde el navegador. Las URL y los nombres de archivo están en `fuentes/README.md`.
- **Tampoco se localizaron** copias de acceso abierto de la Ley N.º 133/93, el Decreto-Ley N.º 46/72 y la Ley N.º 5.254/14. Los tres se citan en el cuerpo. El contenido operativo de la Ley N.º 5.254/14 está, de todos modos, transcripto en los considerandos del Decreto N.º 2129/14, que sí está archivado.
- **Texto literal de la Resolución SDCU N.º 1.670/2022**, citada en el § 2.3. No está en el sitio de la SEDECO; por ahora se archivó la nota oficial de la Agencia IP. Hace falta para resolver si el redondeo hacia arriba que aplica el MRE se ajusta al mandato de que el ajuste sea favorable al consumidor.
- **La Ley N.º 7196/23 ya no se cita en el cuerpo** tras eliminarse el § 2.5, pero sigue en la bibliografía. Mismo criterio pendiente que con la Ley N.º 6935/22: o se quita la entrada, o se acepta la excepción.
- **Ficha del Decreto N.º 6225/2026** (jornal mínimo): la entrada bibliográfica apunta al sitio del MTESS y no al texto del decreto en Gaceta Oficial. Conviene reemplazar la URL por la fuente normativa directa.
- **El jornal mínimo tiene fecha de vencimiento.** El Decreto N.º 6225/2026 rige hasta el 30 de junio de 2027. Si el trabajo se presenta después, el § 2.3 y los importes citados quedan desactualizados.
- **La Ley N.º 6935/22 ya no se cita en el cuerpo** tras la corrección del § 1.3, pero sigue en la bibliografía. APA 7 pide que la lista de referencias sólo contenga obras citadas en el texto. Sostiene la ficha el hecho de que la carátula invoca su artículo 110; si se prefiere el criterio estricto, hay que quitar la entrada.
- El **Capítulo V** debe organizarse en subtítulos explícitos de **viabilidad técnica / operativa / normativa**, conforme a la observación del profesor.
- **`## Objetivos` es un título de nivel 2**, no de nivel 1: así fluye a continuación de la Introducción, que termina anunciándolos, en vez de abrir página nueva. En el índice aparece anidado bajo Introducción. Si se prefiere como sección independiente, cambiar `##` por `#` en `cuerpo/objetivos.md`.
- El script **no procesa imágenes** desde Markdown. La Resolución exige que gráficos, tablas y cuadros se inserten **como imagen**; el plan es agregarlos durante la revisión final en Word, o extender el build cuando haga falta.

---

## 7. Convención de trabajo

### El texto redactado no se reescribe

La prosa de los `.md` ya está acordada y revisada. El trabajo sobre este repositorio es **estructurarla, versionarla y formatearla** — no mejorar la redacción. Es un trabajo académico evaluado por un Tribunal y presentado bajo el nombre del autor: la voz y las decisiones argumentativas son suyas.

Si aparece un problema real —una cita incompleta, una contradicción entre capítulos, una regla de formato que el texto incumple— se señala aparte, o como comentario HTML en el `.md` (que el build descarta), y se deja la decisión al autor.

Las secciones que el autor no redactó (carátula, agradecimientos) sí se pueden draftear, marcándolas como borrador.

### Extensión: de 3 a 4 páginas por capítulo

Decisión del autor. Con cinco capítulos de ese tamaño, más la parte previa —Introducción,
Objetivos, Justificación, Metodología y Limitaciones, que ocupan ocho páginas— y las Conclusiones,
el trabajo supera con holgura el mínimo de 20 páginas sin volverse un tratado.

De ahí se sigue un criterio de redacción: **los capítulos no se extienden en análisis normativo**.
La norma se cita, se dice qué regla establece y se sigue adelante. Las discusiones de vigencia, las
cadenas de derogaciones y los cotejos entre instrumentos quedan fuera del cuerpo; si hacen falta
como respaldo, van en una nota de trabajo del `.md` o en `fuentes/README.md`.

Para medir, lo único que sirve es Word: `verificar_en_word.ps1` informa en qué página física
arranca cada título, y la resta da la extensión real del capítulo. El estimador por palabras del
build aproxima el total del documento, pero no sirve capítulo por capítulo.

Como referencia, el Capítulo I son 1.031 palabras en 3 páginas: unas **340 palabras por página**.

### Un commit por sección

```
Redacción completa del Capítulo II: diagnóstico del procedimiento actual
Corrección de estilo en la Justificación
Completar ficha APA de Alfonso (1995) en la bibliografía
```

Así el historial refleja la evolución del trabajo y permite volver a cualquier versión anterior de una sección.

### Flujo típico de una sesión

1. Pegar o redactar el texto en el `.md` que corresponda.
2. `python build/generar_docx.py`
3. `python build/verificar.py` — si falla algún conteo esperado, revisar si el cambio fue intencional y actualizar la constante.
4. `powershell -ExecutionPolicy Bypass -File build/verificar_en_word.ps1` si interesa el recuento real de páginas.
5. Commit con mensaje descriptivo y `git push`.

---

## 8. Bitácora

Qué se hizo y por qué, para no repetir el análisis.

### Montaje del repositorio

Se armó la estructura completa con el contenido ya redactado (preliminares, introducción, objetivos, justificación, metodología, limitaciones y Capítulo I), más placeholders para los capítulos II a V y las conclusiones. Se escribió el build en `python-docx` puro, descartando Pandoc por las razones de la sección 5.

La carátula se resolvió con un **bloque YAML** al inicio de `preliminares/caratula.md` en vez de texto corrido: la Resolución pide bloques en 24 pt conviviendo con bloques en 12 pt, y eso no se expresa en Markdown. El script lee ese bloque y compone la portada.

El agradecimiento **no venía redactado**; se dejó un borrador marcado como tal, para reemplazar o eliminar.

### Verificación

Se instaló Python 3.12 y se corrió el build de punta a punta, abriendo el resultado en Word para actualizar los campos. En el camino aparecieron y se corrigieron:

- **El índice se listaba a sí mismo** — ver "La trampa del campo TOC" en la sección 5.
- **La estimación de extensión estaba mal calibrada**: `PALABRAS_POR_PAGINA` pasó de 380 (estimado) a 250 (medido sobre el documento ya compuesto).
- Un `FutureWarning` de lxml al resolver el elemento XML en `aplicar_fuente`.

### Mudanza a repositorio propio

El trabajo nació dentro de `esecaese/academia`, en `segundo-semestre/monografia`. Se movió a `esecaese/monografia` con **`git subtree split`**, que conserva la historia: los dos primeros commits son los originales, con sus mensajes y fechas.

Los commits de la monografía en `academia` nunca se habían publicado, y el commit que quitaba la carpeta dejaba el contenido idéntico a `origin/main` — o sea, en neto no cambiaban nada. Se descartaron con `git reset --hard origin/main` para no dejar un rodeo visible en el historial del proyecto principal. En `academia` ya no queda rastro de la monografía.

### Corrección de la base tarifaria de la Apostilla

El Capítulo I afirmaba que ninguna norma fijaba una tasa expresa para la Apostilla y dejaba la
cuestión abierta al Capítulo II. Resultó ser incorrecto. La cadena es: la **Ley N.º 5.254/14**
modificó los artículos 3º, 6º, 12 y 14 de la **Ley N.º 4.033/10** (Arancel Consular) y, en su
artículo 6º in fine, facultó al Poder Ejecutivo y al MRE a fijar los montos; en ejercicio de esa
facultad se dictó el **Decreto N.º 2129/14**, que establece la Apostilla en **dos jornales mínimos
diarios**, designa a la Dirección de Legalizaciones como emisora y perceptora, y exonera del pago
a los beneficiarios del artículo 17 de la Ley N.º 4.033/10, a las instituciones públicas y a los
diplomáticos por reciprocidad.

Se verificó contra el texto oficial: el PDF de SILpy es un escaneo, y se le extrajo la capa de
texto OCR descomprimiendo los streams del PDF con `zlib` (no hay poppler en este equipo).

Efecto colateral: el párrafo del § 1.1 que descartaba la Ley N.º 4.033/10 como base legal pasó a
contradecir al § 1.2, porque la tasa de la Apostilla deriva justamente de esa línea normativa.
Se quitó.

Dos hallazgos que conviene no perder, porque son argumento y no sólo dato:

- La ley habilitante fija como finalidad expresa de la delegación **proceder a través de medios informáticos**. Es respaldo normativo directo para la propuesta, útil en el Capítulo V (viabilidad normativa).
- Las exoneraciones del artículo 3º condicionan el arancel **al sujeto solicitante**, no al documento. Es una regla de decisión que el asistente debe modelar aparte del tipo de trámite (Capítulos III y IV).

### Fecha de la carátula

La portada arrastraba el marcador `[MES] de [AÑO]`, que era una forma segúra de entregar el
trabajo con el marcador puesto. Ahora `mes_anio_24pt` acepta el valor `auto` y el build lo
resuelve al mes y año en que se corre, con los nombres de `MESES` en `formato.py`.

Se dejó como valor y no como código fijo a propósito: escribir la fecha literal en
`caratula.md` sigue funcionando y tiene prioridad sobre `auto`. Eso importa al entregar,
porque una fecha automática cambia sola si el documento se regenera al mes siguiente.

### Títulos en modo oración antes de los capítulos

Indicación del profesor: donde arrancan los números arábigos —la Introducción— el título debe ir
en modo oración y pegado a la izquierda, lo mismo Objetivos con sus dos subtítulos, y **recién al
iniciar los capítulos** vuelve al formato centrado y en mayúsculas. Se extendió el criterio a
Justificación, Metodología y Limitaciones, que quedan entre Objetivos y el Capítulo I y son el
mismo tipo de sección. Conclusiones y Bibliografía quedan como los capítulos.

Dos cosas que hicieron el cambio más chico de lo que parecía:

- **Objetivos no necesitó código.** Es un título de nivel 2, y los niveles 2 y 3 ya van a la izquierda y ya respetan la caja tal como se escribe en el `.md`. Alcanzó con corregir `objetivos.md`, lo que de paso cerró el pendiente de `### GENERAL` / `### ESPECÍFICOS` en mayúsculas.
- **La caja viaja en el `.md`, no en el código.** Los títulos afectados se escriben ya en modo oración (`# Introducción`) y el build deja de aplicarles `.upper()`. Así lo que se lee en el Markdown es lo que sale impreso, igual que en los niveles 2 y 3.

Lo único delicado fue la alineación. La tentación es crear un estilo nuevo, y es la trampa del
campo TOC otra vez: un estilo propio habría sacado esas secciones del índice. Se conserva
`Heading 1` y sólo se sobreescribe la alineación en el párrafo, con formato directo.

`verificar.py` pasó de 52 a 55 comprobaciones: el chequeo "todos en mayúsculas" no podía seguir
valiendo para todos, así que ahora distingue los dos regímenes y verifica, sobre los títulos en
modo oración, que sean exactamente cuatro, que lleven sólo la primera letra en mayúscula y que
estén efectivamente alineados a la izquierda.

### Redacción del Capítulo II

El capítulo se redactó sobre fuentes verificadas, no sobre supuestos. Tres cuestiones quedaron
resueltas en el camino y conviene no volver a abrirlas:

- **El jornal mínimo diario es de G. 117.077** desde el 1 de julio de 2026, por el Decreto N.º 6225/2026 del 17 de junio de 2026, y rige hasta el 30 de junio de 2027.
- **El Decreto N.º 2129/14 sigue vigente.** Se leyó el ejemplar oficial sancionado de la Ley N.º 7196/23 (SILpy): deroga los artículos 8º, 9º, 10, 13, 15, 16 y 14 de la Ley N.º 4.033/10, más capítulos de los artículos 4º y 11. El artículo 6º —la delegación en que el Decreto se funda— no está derogado, y el artículo 17 —al que remiten las exoneraciones— tampoco. De los cuatro artículos que la Ley N.º 5.254/14 había modificado (3º, 6º, 12 y 14), sólo cayó el 14.
- **La tabla de precios del MRE redondea hacia arriba.** Los importes publicados no son el producto exacto del coeficiente por el jornal, sino ese producto redondeado al múltiplo superior de Gs. 50. Se cotejaron los seis coeficientes de la escala (½, 1, 2, 3, 5 y 10 jornales) y los seis coinciden. Ninguna norma establece ese redondeo: es convención administrativa. En la Apostilla representa G. 46 sobre el valor legal exacto.

Dos hallazgos menores que el capítulo aprovecha como argumento:

- El portal de legalizaciones del MRE lista sus instrumentos legales y **omite la Ley N.º 5.254/14 y el Decreto N.º 2129/14**, que son justamente los que fundan el monto de la Apostilla que ese mismo portal cobra. Sirve para mostrar que la dispersión normativa no es un problema teórico.
- El Manual Consular del MRE (abril de 2025) **no sirve para este capítulo**: es del Servicio Exterior, no de la Dirección de Legalizaciones. Cero menciones de "jornal", de la Ley N.º 1.030/97 y del Decreto N.º 2129/14. No volver a descargarlo para esto.

Sobre el método de lectura de los PDF oficiales: el de la Ley N.º 7196/23 es un escaneo **sin capa de
texto**, de modo que la técnica de descomprimir streams con `zlib` —la que sirvió para el Decreto
N.º 2129/14— no da nada. Lo que funcionó fue extraer el JPEG embebido (`/DCTDecode`) del PDF y leer
la imagen directamente. Queda anotado porque es el camino corto para cualquier otro escaneo de SILpy.

**BACN bloquea las descargas automatizadas** (403 y 500 según el caso). SILpy, en cambio, responde
sin problema con un user-agent de navegador, y el sitio del MRE también. Para los textos legales,
ir primero a SILpy.

Dos entradas nuevas de bibliografía (16 → 18): el Decreto N.º 6225/2026 y la página institucional de
legalizaciones del MRE. `ENTRADAS_BIBLIOGRAFIA` y `VINETAS_OBJETIVOS` se actualizaron en
`verificar.py`; esta última ahora cuenta 12, porque el § 2.2 agrega seis viñetas a las seis de
Objetivos. El nombre de la constante quedó corto, pero se conservó para no tocar más de lo necesario.

### Carpeta de fuentes y correcciones al Capítulo II

Se armó `fuentes/` con todo lo que el trabajo cita. En el camino, dos documentos que el autor
aportó —la Ley N.º 4.033/10 y el Decreto N.º 6225/2026— obligaron a corregir el Capítulo II en dos
puntos, uno de ellos importante.

**La delegación estaba mal atribuida.** Se había escrito que el artículo 6º de la Ley N.º 4.033/10
contenía la delegación en que se funda el Decreto N.º 2129/14. Es incorrecto: ese artículo trata de
estampillas y forma de percepción. La delegación está en el **artículo 6º in fine de la Ley
N.º 5.254/14**, según los considerandos del propio Decreto. El apartado que lo discutía terminó
eliminándose (ver más abajo), pero el dato se conserva en la nota de trabajo del capítulo porque
hace falta para el Capítulo V.

**El artículo 17 resultó ser mucho más que un dato que faltaba.** Es el fuero de pobreza, y manda
conceder la exoneración "excepcionalmente en los casos indispensables y con criterio restrictivo",
previa comprobación por los funcionarios consulares. Tres consecuencias:

- Su destinatario es el connacional residente en el exterior, es decir, uno de los colectivos que motivan el trabajo.
- Es el único componente discrecional de todo el régimen: el cálculo del arancel es reglado, pero esta exoneración está deliberadamente confiada al juicio del funcionario.
- Por lo tanto marca el límite material de la automatización, y concuerda con la conclusión de Dávila Elguera (2023) ya citada en el § 1.3.

El § 2.4 se amplió con la cita textual y ese análisis, y el cierre del capítulo pasó de tres
elementos a cuatro. El Capítulo IV tendrá que prever una derivación al funcionario competente en
lugar de resolver esa causal por sí mismo.

Conviene no perder una precisión de método: el artículo 6º in fine de la Ley N.º 5.254/14 fija como
finalidad expresa de la delegación **proceder a través de medios informáticos**. Es respaldo
normativo directo para la propuesta, y el § 2.5 ya lo deja anotado para la viabilidad normativa del
Capítulo V.

Sobre la fecha del Decreto N.º 2129/14: el OCR del ejemplar oficial confirma el día 26 y el año
2014, pero el mes sigue ilegible. Agosto continúa siendo lo más probable y el pendiente sigue
abierto.

### Recorte del Capítulo II y base normativa del redondeo

Tres decisiones del autor, todas en la misma dirección: que los capítulos no se vuelvan un tratado
de derecho.

**El redondeo sí está reglamentado.** El § 2.3 lo presentaba como una convención administrativa sin
formulación expresa. Es incorrecto: la **Resolución SDCU N.º 1.670/2022** de la SEDECO, que abrogó
la Resolución N.º 347/14, obliga a fijar los precios en cifras redondas y a practicar el ajuste en
la denominación de cincuenta guaraníes, alcanzando expresamente a los comprobantes de los servicios
públicos. El apartado se reescribió: el importe percibido resulta de encadenar dos reglas de fuente
distinta —la escala arancelaria, que da el coeficiente, y la normativa de redondeo monetario, que
fija la expresión final—, y eso es justamente lo que una automatización debe explicitar. La
verificación aritmética se conserva, porque sigue probando que la tabla publicada se deriva del
jornal vigente.

**Se eliminaron dos apartados.** El de la vigencia tras la Ley N.º 7196/23 y el de la dispersión
normativa. Razón del autor: la derogación de 2023 no incide en la práctica sobre las apostillas y
legalizaciones en que la Dirección de Legalizaciones tiene rol local, y el capítulo no debe
extenderse en análisis de leyes. Son 622 palabras menos. El capítulo quedó en cinco apartados y el
cuerpo computable sigue en 21 páginas, por encima del mínimo.

Dos hallazgos de los apartados eliminados quedaron guardados en la nota de trabajo del
`capitulo-2.md`, porque no conviene perderlos:

- La base legal de la tasa está intacta: la Ley N.º 7196/23 no toca ninguna disposición de la Ley N.º 5.254/14, donde vive la delegación.
- Esa delegación fija como finalidad expresa **proceder a través de medios informáticos**. Es respaldo normativo directo para la propuesta y corresponde usarlo en la viabilidad normativa del Capítulo V.

**Efecto colateral en la bibliografía.** La Ley N.º 7196/23 ya no se cita en el cuerpo, pero sigue
en la lista de referencias. Es el mismo caso que la Ley N.º 6935/22: APA 7 pide que la lista
contenga sólo obras citadas. Queda como pendiente de decisión.

Una cuestión que el § 2.3 deliberadamente no aborda: la Resolución SDCU N.º 1.670/2022 manda que el
ajuste sea **favorable al consumidor**, y la tabla del MRE redondea siempre hacia arriba —G. 46 de
más en la Apostilla—, o sea en contra del usuario. Se dejó fuera para no alargar el capítulo, pero
está anotado en el `capitulo-2.md` por si conviene retomarlo en el Capítulo V. Antes de afirmarlo
hace falta el texto literal de la Resolución, que todavía no se consiguió.

### Rastrillos ya pisados

- `verificar_en_word.ps1` tuvo que guardarse **con BOM UTF-8**: Windows PowerShell 5.1 lee los `.ps1` sin BOM como ANSI y destroza los acentos de los mensajes.
- `Documento.SaveAs()` por COM exige `[ref]` sobre **variables tipadas**; pasarle directamente el resultado de `Join-Path` falla con "no se puede convertir el valor de tipo psobject".
- Correr `verificar_en_word.ps1` dos veces seguidas muy rápido falla con `COMException`: la instancia anterior de Word sigue cerrándose.

---

## 8 bis. Las fuentes

`fuentes/` guarda una copia local de todo lo que el trabajo cita o consulta: leyes, decretos,
papers, informes, páginas institucionales y extractos. El índice comentado está en
[`fuentes/README.md`](fuentes/README.md), con qué aporta cada documento y a qué apartado sirve.

La razón de versionarlas es práctica: varias normas viven en portales que cambian de URL, bloquean
descargas automatizadas o directamente caen. El trabajo tiene que poder defenderse ante el Tribunal
sin depender de que un sitio oficial siga en línea el día de la presentación.

**Los libros con derechos de autor no se versionan.** Hernández Sampieri y otros (2014) y Russell y
Norvig (2021) son obras comerciales y este repositorio es público: subirlas sería redistribuirlas.
Para tener copias a mano mientras se trabaja está `fuentes/_privado/`, que el `.gitignore` excluye.

### Leer los PDF oficiales escaneados

En este equipo no hay poppler ni OCR, así que para los ejemplares escaneados sirven dos caminos:

- **Con capa OCR** (el Decreto N.º 2129/14): descomprimir los streams con `zlib` y juntar los literales entre paréntesis. Los acentos vienen en octal y hay que decodificarlos.
- **Sin capa OCR** (la Ley N.º 7196/23): extraer el JPEG embebido —el objeto `/Subtype /Image` con filtro `/DCTDecode`— y leer la imagen directamente.

El Decreto N.º 6225/2026 es un tercer caso: tiene texto real, pero con una fuente CID subsetada que
no se decodifica limpiamente sin poppler. Sus cifras se confirmaron contra el sitio del MTESS.

---

## 9. Repositorio

Repositorio propio, separado de las materias del curso: [`esecaese/monografia`](https://github.com/esecaese/monografia). Es la copia de referencia: el proyecto se clona en cada equipo donde se trabaja y no vive en una ruta fija. Con el remoto ya configurado, para publicar cambios alcanza con:

```bash
git push
```

---

## 10. Arrancar una sesión en cualquier equipo

El proyecto no depende de una máquina en particular: los `.md` son la fuente de verdad, el `.docx`
se regenera, y el build corre en cualquier sistema con Python. Para retomar el trabajo en otro
equipo alcanza con clonar el repositorio y pegar este prompt:

```text
Estoy retomando mi trabajo monográfico para la Academia Diplomática y Consular
"Carlos Antonio López". El proyecto vive en GitHub:

    https://github.com/esecaese/monografia

Arrancá así:

1. Si no está clonado en este equipo, clonalo. Si ya está, hacé git pull para
   traer lo último.

2. Leé el README.md completo antes de hacer nada. Es autosuficiente: tiene el
   contexto, el formato exigido por la Resolución N.º 1055/2024, el estado de
   avance, la bitácora de decisiones ya tomadas y las convenciones de trabajo.

3. Configurá la identidad de git local al repositorio. No tengo identidad
   global, así que sin esto los commits fallan:

       git config --local user.name "esecaese"
       git config --local user.email "knut.sach@gmail.com"

   Si el push pide autenticación, resolvelo con gh auth login.

4. Instalá la dependencia: pip install -r build/requirements.txt
   Es la única (python-docx). No hace falta Pandoc. Si el comando "python" no
   responde, buscá el intérprete en este equipo antes de darte por vencido —
   suele no estar en el PATH.

5. Corré el build y la verificación para confirmar que todo sigue sano, y
   decime en qué estado está el trabajo:

       python build/generar_docx.py
       python build/verificar.py

Dos cosas que quiero que tengas presentes desde el arranque:

1. El texto de los .md ya está redactado y acordado. No lo reescribas ni lo
   "mejores" — solo estructurá y formateá. Si ves un problema real,
   señalamelo aparte y decido yo.

2. Un commit por sección o capítulo, con mensaje descriptivo, y push al final.

Hoy quiero trabajar en el Capítulo II (diagnóstico del procedimiento actual de
liquidación). Te voy a ir pegando el texto para que lo integres y versiones.
```

La última línea es la que cambia según la sesión; el resto queda igual siempre.

El prompt puede ser tan corto porque este README hace el trabajo pesado: apuntar acá y decir
"leelo" alcanza. Por eso conviene mantenerlo al día.

### Lo que cambia según el equipo

- **La identidad de git no se hereda.** Está configurada local al repositorio y no a nivel global, así que un clon nuevo viene sin ella y los commits fallan. Por eso el prompt la incluye como paso explícito.
- **El push es por HTTPS** y pide credenciales en cada equipo nuevo. `gh auth login` lo resuelve.
- **`verificar_en_word.ps1` sólo corre en Windows con Word instalado.** Queda deliberadamente fuera del prompt para que no se intente donde va a fallar. Donde haya Word se pide aparte: es el único que da el recuento **real** de páginas contra el mínimo de 20, que el estimador del build sólo aproxima —y por debajo—.
- **`output/` está en `.gitignore`.** El `.docx` no viaja por git: se regenera en cada equipo con `python build/generar_docx.py`.
- **Antes de borrar una copia local**, comprobar que no quede nada sin publicar: `git status` limpio y `git log origin/main..HEAD` vacío. Y cerrar Word, porque el archivo de bloqueo `~$*.docx` impide borrar la carpeta.
