<#
Diagnostico pontual (31/08/2026): a extracao normal so loga progresso DEPOIS que uma aba inteira
termina de exportar - se travar antes disso (na abertura do arquivo, por exemplo) nao aparece
nada no log, so o timeout de 15min. Este script loga cada passo separadamente pra achar
exatamente onde trava, sem precisar esperar o timeout inteiro.
#>

$ExcelPath = "C:\Users\Daterrinha63\daterrinhaalimentos.com.br\DT - COMEX - Fluxo de Exportação\Fluxo_Exportações_2026.xlsx"

function Log { param([string]$m) Write-Host ("[{0}] {1}" -f (Get-Date -Format "HH:mm:ss.fff"), $m) }

Log "INICIO do diagnostico"
Log "Criando objeto Excel.Application..."
$excel = New-Object -ComObject Excel.Application
Log "OK - Excel.Application criado"

$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
$excel.EnableEvents = $false
Log "Propriedades basicas configuradas (Visible/DisplayAlerts/AskToUpdateLinks/EnableEvents)"

$workbook = $null
try {
    Log "Abrindo workbook (ReadOnly, UpdateLinks=0)..."
    $workbook = $excel.Workbooks.Open($ExcelPath, 0, $true)  # UpdateLinks=0 (nao atualiza links externos), ReadOnly=true
    Log "OK - Workbook aberto"

    Log ("Numero de conexoes de dados no workbook: " + $workbook.Connections.Count)
    Log ("Numero de abas: " + $workbook.Sheets.Count)

    Log "Listando nomes das abas..."
    foreach ($sheet in $workbook.Sheets) {
        Log ("  aba: '" + $sheet.Name + "'")
    }

    Log "Tentando ler uma unica celula da primeira aba (teste de responsividade)..."
    $firstSheet = $workbook.Sheets.Item(1)
    $val = $firstSheet.Cells.Item(1,1).Text
    Log ("OK - celula A1 da primeira aba leu: '" + $val + "'")
}
finally {
    Log "Fechando workbook..."
    if ($workbook) { $workbook.Close($false) }
    Log "Encerrando Excel..."
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Log "FIM do diagnostico"
}
