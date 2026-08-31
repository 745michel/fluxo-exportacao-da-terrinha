<#
Abre Fluxo_Exportações_2026.xlsx (Excel COM, somente leitura) e extrai as 3 abas de
programação de exportação (2024/2025/2026) para CSV em data/raw/, usando Cell.Text (texto
formatado, não Value2) para que os números venham em "1.234,56" e as datas em "dd/mm/aaaa" ou
"dd-mmm-aa" exatamente como o build.js (parseDateValue/toNumber) espera.

Extrai também a coluna "Código Produto" da aba 2026 separadamente (codigo_2026.csv), porque
essa coluna tem células com erro de fórmula (#N/A etc.) que .Text detecta e o build.js filtra.
#>

param(
    [string]$ExcelPath = "C:\Users\Daterrinha63\daterrinhaalimentos.com.br\DT - COMEX - Fluxo de Exportação\Fluxo_Exportações_2026.xlsx"
)

$ErrorActionPreference = "Stop"

$rawDir = Join-Path (Split-Path -Parent $PSScriptRoot) "data\raw"
New-Item -ItemType Directory -Force -Path $rawDir | Out-Null

if (-not (Test-Path -LiteralPath $ExcelPath)) {
    throw "Planilha nao encontrada: $ExcelPath"
}

function Escape-CsvField {
    param([string]$Value)
    if ($null -eq $Value) { return '' }
    if ($Value -match '[",\r\n]') {
        return '"' + ($Value -replace '"', '""') + '"'
    }
    return $Value
}

# UsedRange.Rows.Count NAO e confiavel: uma formatacao acidental numa coluna inteira (ex.: alguem
# selecionou a coluna toda e pintou/formatou sem querer) faz o Excel "achar" que a planilha tem
# dado ate a ULTIMA linha possivel (1.048.576) mesmo com os dados reais só nas primeiras centenas
# de linhas - descoberto em 31/08/2026 na aba de 2024 (38 milhoes de celulas "usadas"), o que
# fazia a extracao celula-por-celula travar por horas sem nunca terminar. Cells.End(xlUp) a partir
# da ultima linha do sheet e uma chamada COM so (rapida, o Excel resolve internamente), não um
# loop célula a célula - acha a última linha com dado de verdade na coluna informada.
$XL_UP = -4162
function Get-RealLastRow {
    param($Sheet, [int]$StartRow, [int]$Col)
    $bottomCell = $Sheet.Cells.Item($Sheet.Rows.Count, $Col)
    $lastRow = $bottomCell.End($XL_UP).Row
    if ($lastRow -lt $StartRow) { return $StartRow }
    return $lastRow
}

function Export-SheetToCsv {
    param($Sheet, [string]$OutPath)
    # Range.Text só é válido célula a célula via COM (não existe em bloco para um range com mais
    # de uma célula) - por isso a leitura é feita uma célula de cada vez, não via array 2D.
    $used = $Sheet.UsedRange
    $startRow = $used.Row
    $startCol = $used.Column
    $cols = $used.Columns.Count
    $realLastRow = Get-RealLastRow -Sheet $Sheet -StartRow $startRow -Col $startCol
    $rows = $realLastRow - $startRow + 1
    if ($rows -ne $used.Rows.Count) {
        Write-Host ("AVISO: UsedRange da aba '" + $Sheet.Name + "' dizia " + $used.Rows.Count + " linhas, mas a ultima linha com dado de verdade e a " + $realLastRow + " - usando " + $rows + " linhas reais.")
    }
    $cellsObj = $Sheet.Cells

    $sb = New-Object System.Text.StringBuilder
    for ($r = 0; $r -lt $rows; $r++) {
        $fields = New-Object System.Collections.Generic.List[string]
        for ($c = 0; $c -lt $cols; $c++) {
            $val = $cellsObj.Item($startRow + $r, $startCol + $c).Text
            $fields.Add((Escape-CsvField $val))
        }
        [void]$sb.AppendLine(($fields -join ','))
    }
    [System.IO.File]::WriteAllText($OutPath, $sb.ToString(), [System.Text.Encoding]::UTF8)
}

function Export-ColumnToCsv {
    param($Sheet, [int]$ColIndex, [string]$OutPath)
    $used = $Sheet.UsedRange
    $startRow = $used.Row
    $startCol = $used.Column
    $realLastRow = Get-RealLastRow -Sheet $Sheet -StartRow $startRow -Col $startCol
    $rows = $realLastRow - $startRow + 1
    $sb = New-Object System.Text.StringBuilder
    for ($r = 0; $r -lt $rows; $r++) {
        $cell = $Sheet.Cells.Item($startRow + $r, $ColIndex)
        [void]$sb.AppendLine($r.ToString() + ',' + (Escape-CsvField $cell.Text))
    }
    [System.IO.File]::WriteAllText($OutPath, $sb.ToString(), [System.Text.Encoding]::UTF8)
}

# Busca a aba pelo nome ignorando espacos duplicados/extras - a aba de 2024 mudou de "Programação
# Exportações 2024" pra "Programação  Exportações 2024" (dois espacos) numa edicao recente da
# planilha (28/08/2026), e o Item($nome) exato quebrou com "Indice invalido". Comparar normalizado
# evita quebrar de novo se isso mudar de novo no futuro.
function Find-SheetByNormalizedName {
    param($Workbook, [string]$TargetName)
    $normalizedTarget = ($TargetName -replace '\s+', ' ').Trim()
    foreach ($sheet in $Workbook.Sheets) {
        $normalizedActual = ($sheet.Name -replace '\s+', ' ').Trim()
        if ($normalizedActual -eq $normalizedTarget) { return $sheet }
    }
    return $null
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $null
try {
    $workbook = $excel.Workbooks.Open($ExcelPath, [Type]::Missing, $true)  # ReadOnly = $true

    $sheetMap = @{
        2024 = "Programação Exportações 2024"
        2025 = "Programação Exportações 2025"
        2026 = "Programação Exportações 2026"
    }

    $sheets = @{}
    foreach ($year in $sheetMap.Keys) {
        $sheetName = $sheetMap[$year]
        $sheet = Find-SheetByNormalizedName -Workbook $workbook -TargetName $sheetName
        if (-not $sheet) { throw "Aba nao encontrada (nem normalizando espacos): '$sheetName'" }
        $sheets[$year] = $sheet
        $outPath = Join-Path $rawDir "Programação_Exportações_$year.csv"
        Export-SheetToCsv -Sheet $sheet -OutPath $outPath
        Write-Host "Exportado: $($sheet.Name) -> $outPath"
    }

    # Coluna "Código" da aba 2026: offset fixo 6 (0-based) a partir da primeira coluna usada -
    # mesmo indice que build.js espera em SCHEMAS[2026].codigo. Confirmado contra o cabecalho
    # real (Cliente=0, Tipo=1, Pedido=2, Invoice=3, Data=4, Volume=5, Codigo=6, Descricao=7, ...).
    # Extraida separada da aba inteira porque so aqui .Text detecta as celulas com erro de
    # formula (#N/A etc.) que o build.js precisa descartar.
    $sheet2026 = $sheets[2026]
    $used = $sheet2026.UsedRange
    $codigoCol = $used.Column + 6
    Export-ColumnToCsv -Sheet $sheet2026 -ColIndex $codigoCol -OutPath (Join-Path $rawDir "codigo_2026.csv")
    Write-Host "Exportado: coluna Codigo (2026) -> codigo_2026.csv"
}
finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
