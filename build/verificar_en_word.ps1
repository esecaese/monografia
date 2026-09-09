<#
    Verificación en Word del .docx generado.

        powershell -ExecutionPolicy Bypass -File build\verificar_en_word.ps1

    Lo que build\verificar.py no puede hacer: abrir el documento en Word,
    actualizar los campos de verdad y preguntarle a Word cuántas páginas
    salieron y con qué numeración. Sirve para confirmar que el índice se
    arma bien y para medir la extensión real contra el mínimo de 20 páginas
    que exige la Resolución N.º 1055/2024.

    Abre el archivo, actualiza, informa, exporta un PDF a output\ y cierra
    SIN guardar: el .docx del repositorio queda intacto.

    Requiere Microsoft Word instalado.
#>

param(
    [string]$Docx = "output\monografia-aranceles-consulares.docx",
    [string]$Pdf  = "output\vista-previa.pdf"
)

$ErrorActionPreference = "Stop"
$raiz = Split-Path -Parent $PSScriptRoot
$rutaDocx = Join-Path $raiz $Docx
$rutaPdf = Join-Path $raiz $Pdf

if (-not (Test-Path $rutaDocx)) {
    throw "No existe $rutaDocx. Correr antes: python build\generar_docx.py"
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    $doc = $word.Documents.Open($rutaDocx, $false, $false)
    $doc.Fields.Update() | Out-Null
    foreach ($toc in $doc.TablesOfContents) { $toc.Update() }
    $doc.Repaginate()

    # wdStatisticPages = 2, wdStatisticWords = 0
    $paginas = $doc.ComputeStatistics(2)
    "Páginas totales : $paginas"
    "Palabras totales: " + $doc.ComputeStatistics(0)
    "Secciones       : " + $doc.Sections.Count

    "`n--- Índice generado ---"
    if ($doc.TablesOfContents.Count -gt 0) {
        $doc.TablesOfContents.Item(1).Range.Text -replace "`r", "`n"
    } else {
        "SIN ÍNDICE: el campo TOC no se encontró."
    }

    "`n--- Página física donde arranca cada título de nivel 1 ---"
    foreach ($p in $doc.Paragraphs) {
        if ($p.OutlineLevel -eq 1 -and $p.Range.Text.Trim().Length -gt 1) {
            # wdActiveEndPageNumber = 3
            "{0,4}  {1}" -f $p.Range.Information(3), $p.Range.Text.Trim()
        }
    }

    # wdFormatPDF = 17. SaveAs exige [ref] a variables tipadas: pasarle
    # directamente el resultado de Join-Path falla con "no se puede convertir
    # el valor de tipo psobject".
    [string]$destino = $rutaPdf
    [int]$formatoPdf = 17
    $doc.SaveAs([ref]$destino, [ref]$formatoPdf)
    "`nVista previa exportada: $Pdf"

    # wdDoNotSaveChanges = 0
    $doc.Close([ref]0)
}
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
