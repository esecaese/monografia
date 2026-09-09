# Monografía — Automatización de la liquidación de aranceles consulares mediante inteligencia artificial

Trabajo monográfico del **Curso de Perfeccionamiento** de la Academia Diplomática y Consular "Carlos Antonio López" (Ministerio de Relaciones Exteriores del Paraguay), presentado en el marco del artículo 110 de la Ley N.º 6935/22.

- **Alumno:** Sven Knutson Sachelaridi
- **Categoría en el Escalafón:** Segundo Oficial en Administración y Asuntos Técnicos
- **Dependencia:** Dirección de Informática
- **Correo:** sknutson@mre.gov.py

**Tema.** Propuesta de un asistente virtual integrado al portal de apostillas que evacúe consultas, guíe al usuario y calcule automáticamente el arancel de legalización/apostilla conforme a la Ley N.º 1.030/97 y la Ley N.º 4.987/13.

**Pregunta de investigación.** ¿En qué medida es viable automatizar la liquidación de aranceles consulares mediante un asistente virtual integrado al sistema de apostilla del MRE, bajo el marco normativo vigente y los principios de la gestión digital pública?

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
│   ├── verificar.py             # 52 comprobaciones sobre el .docx generado
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
| `# Título` | nivel 1: mayúsculas, negrita, centrado, página nueva |
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

52 comprobaciones: secciones y `pgNumType`, medidas de hoja y márgenes, estilos (fuentes, tamaños, interlineado, alineaciones), presencia del campo TOC y de `updateFields`, orden y caja de los títulos, y contenido (que los comentarios HTML no se hayan filtrado, que las cursivas se hayan convertido, cantidad de entradas de bibliografía y de viñetas). Devuelve 0 si todo pasa.

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
| Parte pre textual en números romanos, centrados arriba | secciones 1-2, `numerar_paginas(..., "lowerRoman")` |
| Cuerpo en números arábigos desde la Introducción, centrados arriba | sección 3, `numerar_paginas(..., "decimal")` |
| Encabezados de nivel 1 centrados, en mayúscula, negrita, sin punto final | `agregar_titulo` (nivel 1) + estilo `Heading 1` |
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

**Extensión actual: 23 páginas totales, 16 de cuerpo computable** según Word (de la Introducción a las Conclusiones), sobre un mínimo de 20. Buena parte son páginas casi vacías de los capítulos pendientes.

### Pendientes abiertos

- **Mes y año** de presentación en `preliminares/caratula.md` (hoy dice `[MES] de [AÑO]`).
- **Ficha APA completa de Alfonso (1995)**, citado en Metodología y todavía sin referencia localizada.
- **Pregunta operativa heredada del Capítulo I:** bajo qué criterio tarifario se cobra hoy la Apostilla en la práctica, dado que ni la Ley N.º 4.987/13 ni el Decreto N.º 520/13 fijan una tasa expresa. Se responde en el Capítulo II.
- El **Capítulo V** debe organizarse en subtítulos explícitos de **viabilidad técnica / operativa / normativa**, conforme a la observación del profesor.
- **`## OBJETIVOS` es un título de nivel 2**, no de nivel 1: así fluye a continuación de la Introducción, que termina anunciándolos, en vez de abrir página nueva. En el índice aparece anidado bajo INTRODUCCIÓN. Si se prefiere como sección independiente, cambiar `##` por `#` en `cuerpo/objetivos.md`.
- **`### GENERAL` y `### ESPECÍFICOS` están en mayúsculas** en `cuerpo/objetivos.md` y así aparecen en el índice. La Resolución pide que los subtítulos lleven "primera letra mayúscula y el resto minúscula". El texto se dejó tal como fue redactado; para cumplir la regla al pie de la letra habría que escribirlos `### General` y `### Específicos`.
- El script **no procesa imágenes** desde Markdown. La Resolución exige que gráficos, tablas y cuadros se inserten **como imagen**; el plan es agregarlos durante la revisión final en Word, o extender el build cuando haga falta.

---

## 7. Convención de trabajo

### El texto redactado no se reescribe

La prosa de los `.md` ya está acordada y revisada. El trabajo sobre este repositorio es **estructurarla, versionarla y formatearla** — no mejorar la redacción. Es un trabajo académico evaluado por un Tribunal y presentado bajo el nombre del autor: la voz y las decisiones argumentativas son suyas.

Si aparece un problema real —una cita incompleta, una contradicción entre capítulos, una regla de formato que el texto incumple— se señala aparte, o como comentario HTML en el `.md` (que el build descarta), y se deja la decisión al autor.

Las secciones que el autor no redactó (carátula, agradecimientos) sí se pueden draftear, marcándolas como borrador.

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

### Rastrillos ya pisados

- `verificar_en_word.ps1` tuvo que guardarse **con BOM UTF-8**: Windows PowerShell 5.1 lee los `.ps1` sin BOM como ANSI y destroza los acentos de los mensajes.
- `Documento.SaveAs()` por COM exige `[ref]` sobre **variables tipadas**; pasarle directamente el resultado de `Join-Path` falla con "no se puede convertir el valor de tipo psobject".
- Correr `verificar_en_word.ps1` dos veces seguidas muy rápido falla con `COMException`: la instancia anterior de Word sigue cerrándose.

---

## 9. Repositorio

Repositorio propio, separado de las materias del curso: [`esecaese/monografia`](https://github.com/esecaese/monografia). Vive en `C:\Users\svenp\Desktop\Monografia` y el remoto ya está configurado, así que para publicar cambios alcanza con:

```bash
git push
```
