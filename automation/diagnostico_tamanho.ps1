$ExcelPath = "C:\Users\Daterrinha63\daterrinhaalimentos.com.br\DT - COMEX - Fluxo de Exportação\Fluxo_Exportações_2026.xlsx"
function Log { param([string]$m) Write-Host ("[{0}] {1}" -f (Get-Date -Format "HH:mm:ss.fff"), $m) }

Log "Abrindo workbook..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
try {
    $workbook = $excel.Workbooks.Open($ExcelPath, 0, $true)
    Log "OK - Workbook aberto"

    $names = @('Programação Exportações 2024', 'Programação Exportações 2025', 'Programação Exportações 2026')
    foreach ($name in $names) {
        $normalizedTarget = ($name -replace '\s+', ' ').Trim()
        $found = $null
        foreach ($sheet in $workbook.Sheets) {
            $normalizedActual = ($sheet.Name -replace '\s+', ' ').Trim()
            if ($normalizedActual -eq $normalizedTarget) { $found = $sheet; break }
        }
        if ($found) {
            $used = $found.UsedRange
            Log ("Aba '" + $found.Name + "': " + $used.Rows.Count + " linhas x " + $used.Columns.Count + " colunas = " + ($used.Rows.Count * $used.Columns.Count) + " celulas")
        } else {
            Log ("Aba '$name' NAO ENCONTRADA")
        }
    }
    $workbook.Close($false)
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Log "FIM"
}
