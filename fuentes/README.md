# Fuentes

Copia local de todo lo que el trabajo cita o consulta. La idea es que la monografía se pueda
defender sin depender de que un sitio oficial siga en línea: varios de estos documentos están
publicados en portales que cambian de URL, bloquean descargas automatizadas o directamente caen.

Cada archivo lleva el nombre de la norma o del autor, no el nombre con que vino del servidor.

```
fuentes/
├── normativa/        leyes y decretos
├── academicas/       papers, informes y libros
├── institucionales/  páginas y manuales del MRE y otros organismos
└── capturas/         extractos de páginas web que no sobreviven como PDF
```

---

## normativa/

| Archivo | Norma | Origen |
|---|---|---|
| `Ley-1030-1997_Tasas-de-Legalizaciones.pdf` | Ley N.º 1.030/97, escala de tasas de legalizaciones | MRE |
| `Ley-4033-2010_Arancel-Consular.pdf` | Ley N.º 4.033/10, del Arancel Consular | aportado por el autor |
| `Decreto-520-2013_Autoridad-Apostilla.pdf` | Decreto N.º 520/13, designa al MRE autoridad de Apostilla | MRE |
| `Ley-4987-2013_Convenio-Apostilla.pdf` | Ley N.º 4.987/13, aprueba el Convenio de La Haya | SILpy |
| `Decreto-2129-2014_Tasa-Apostilla.pdf` | Decreto N.º 2129/14, fija la tasa de la Apostilla | SILpy |
| `Decreto-6225-2026_Salario-Minimo.pdf` | Decreto N.º 6225/2026, jornal mínimo vigente | aportado por el autor |
| `Ley-7196-2023_Deroga-Arancel-Consular.pdf` | Ley N.º 7196/23, deroga artículos de la Ley N.º 4.033/10 | SILpy |

### Lo que cada una aporta

- **Ley N.º 1.030/97.** El artículo 1º trae la escala completa en jornales mínimos diarios, que es la base del § 2.2. Coeficientes: ½, 1, 2, 3, 5 y 10.
- **Ley N.º 4.033/10.** El **artículo 17** es el fuero de pobreza, citado textual en el § 2.4; manda conceder la exoneración "excepcionalmente en los casos indispensables y con criterio restrictivo". El artículo 18 limita las exoneraciones a las previstas en la ley y en leyes especiales expresas. Ojo: el artículo 6º de **esta** ley trata de estampillas y forma de percepción, no de la delegación.
- **Decreto N.º 2129/14.** Además de fijar la Apostilla en dos jornales, sus considerandos identifican la base legal: el **artículo 6º in fine de la Ley N.º 5.254/14**. Es la fuente que cita el § 2.5. La fecha del ejemplar escaneado trae el mes ilegible ("Asunción, 26 de … de 2014").
- **Ley N.º 7196/23.** Deroga los artículos 8º, 9º, 10, 13, 15, 16 y 14 de la Ley N.º 4.033/10, más capítulos de los artículos 4º y 11. No toca ninguna disposición de la Ley N.º 5.254/14.
- **Decreto N.º 6225/2026.** Jornal mínimo de G. 117.077 desde el 1 de julio de 2026, vigente hasta el 30 de junio de 2027.
- **Resolución SDCU N.º 1.670/2022** (SEDECO). Base normativa del redondeo del § 2.3: obliga a fijar precios en cifras redondas y a ajustar en la denominación de cincuenta guaraníes, alcanzando a los comprobantes de servicios públicos. Abrogó la Resolución N.º 347/14. **Falta el texto literal**: no está publicada en el sitio de la SEDECO, y por ahora sólo se archivó la nota de la Agencia IP. Importa conseguirla porque dispone que el ajuste sea favorable al consumidor, mientras que la tabla del MRE redondea hacia arriba.

---

## academicas/

| Archivo | Referencia |
|---|---|
| `Davila-Elguera-2023_IA-y-tramites-consulares.pdf` | Dávila Elguera (2023), *Política Internacional*, (134), 44-58 |
| `OCDE-BID-2024_Indice-Gobierno-Digital-ALC-2023.pdf` | OCDE/BID (2024), Índice de Gobierno Digital de ALC |

---

## institucionales/

| Archivo | Qué es |
|---|---|
| `MRE_Legalizaciones-Apostilla_2026-10-05.html` | Página oficial del trámite, capturada el 5 de octubre de 2026 |
| `MTESS_Reajuste-salario-minimo_2026-10-05.html` | Comunicado del MTESS sobre el Decreto N.º 6225/2026 |
| `MRE-2025_Manual-Consular.pdf` | Manual Consular, abril de 2025 |
| `SEDECO_Resolucion-1670-2022_nota-AgenciaIP.html` | Nota oficial de la Agencia IP sobre la Resolución SDCU N.º 1.670/2022 |

La página del MRE es la fuente del § 2.1 (sedes, horarios, cadena de intervenciones), del § 2.3
(tabla de precios) y del § 2.6 (la lista de instrumentos legales que omite la Ley N.º 5.254/14 y el
Decreto N.º 2129/14).

El **Manual Consular no sirve para el Capítulo II**: es del Servicio Exterior, no de la Dirección de
Legalizaciones. Cero menciones de "jornal", de la Ley N.º 1.030/97 y del Decreto N.º 2129/14. Queda
archivado porque puede servir para los capítulos de diseño, y para no volver a descargarlo con la
misma esperanza.

---

## capturas/

| Archivo | Qué es |
|---|---|
| `MRE_tabla-arancelaria_2026-10-05.txt` | Tabla de precios del MRE más el cotejo contra la escala legal |

Es la prueba documental del § 2.3: los 25 importes publicados, y la verificación de que los seis
coeficientes de la escala coinciden exactamente con el valor legal redondeado al múltiplo superior
de Gs. 50. Se regenera con el script del mismo nombre si la tabla cambia.

---

## Pendientes de descarga

Dos fuentes citadas que no se pudieron bajar: los dos sitios rechazan las descargas automatizadas.
Hay que guardarlas a mano desde el navegador, con estos nombres, en `academicas/`:

- `Roseth-Reyes-Santiso-2018_El-fin-del-tramite-eterno_BID.pdf` — https://publications.iadb.org/es/publications/spanish/viewer/El-fin-del-tr%C3%A1mite-eterno-Ciudadanos-burocracia-y-gobierno-digital.pdf
- `UNESCO-2021_Recomendacion-etica-IA.pdf` — https://unesdoc.unesco.org/ark:/48223/pf0000381137_spa
- `Resolucion-SDCU-1670-2022_Redondeo.pdf` — texto literal, en `normativa/`. No localizado en línea; probablemente haya que pedirlo a la SEDECO o buscarlo en Gaceta Oficial (publicada el 22 de noviembre de 2022).

Tampoco están, porque no se localizó una copia de acceso abierto:

- **Ley N.º 133/93** y **Decreto-Ley N.º 46/72**, citados en el § 1.1 como antecedentes.
- **Ley N.º 5.254/14**. Su contenido operativo está, de todos modos, transcripto en los considerandos del Decreto N.º 2129/14, que sí está archivado.
- **Alfonso (1995)**, citado en Metodología y todavía sin ficha completa.

---

## Libros con derechos de autor

**No se versionan en este repositorio.** `Hernández Sampieri y otros (2014)` y `Russell y Norvig
(2021)` son obras comerciales, y el repositorio es público: subirlas sería redistribuirlas.

Si querés tener tus copias a mano mientras trabajás, poné los archivos en
`fuentes/_privado/`, que está en el `.gitignore` y no se sube.

---

## Cómo leer los PDF escaneados

Varios ejemplares oficiales son escaneos y no se dejan copiar el texto. En este equipo no hay
poppler ni OCR, así que los caminos que funcionaron son dos:

1. **Si el PDF tiene capa OCR** (el Decreto N.º 2129/14 la tiene), descomprimir los streams con `zlib` y juntar los literales entre paréntesis. Los acentos vienen en octal, hay que decodificarlos.
2. **Si no la tiene** (la Ley N.º 7196/23 no la tiene), extraer el JPEG embebido —el objeto `/Subtype /Image` con filtro `/DCTDecode`— y leer la imagen.

El Decreto N.º 6225/2026 es un caso aparte: tiene texto real, pero con una fuente CID subsetada
cuyo mapa no se decodifica limpiamente sin poppler. Las cifras se confirmaron contra el MTESS.

**BACN rechaza las descargas automatizadas** (403 y 500) y tampoco carga en el navegador.
Para los textos legales conviene ir primero a SILpy, que responde sin problema.
