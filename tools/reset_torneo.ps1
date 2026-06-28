param(
    [string]$ManagerPath
)

$ErrorActionPreference = "Stop"

$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$manager = if ($ManagerPath) {
    (Resolve-Path $ManagerPath).Path
} else {
    Join-Path $root "12HBosaro_manager.xlsx"
}
$backupDirectory = Join-Path $root "assets\backups"

Write-Host ""
Write-Host "RESET SICURO 12 ORE BOSARO" -ForegroundColor Red
Write-Host "Saranno cancellati risultati, squadre knockout e marcatori."
Write-Host "Squadre, gironi, rose, campi e orari resteranno invariati."
Write-Host ""

$confirmation = Read-Host 'Per continuare scrivi esattamente RESET TORNEO'
if ($confirmation -cne "RESET TORNEO") {
    Write-Host "Operazione annullata. Nessun dato modificato." -ForegroundColor Yellow
    exit 0
}

New-Item -ItemType Directory -Path $backupDirectory -Force | Out-Null
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backup = Join-Path $backupDirectory "12HBosaro_manager_prima_reset_$timestamp.xlsx"
Copy-Item -LiteralPath $manager -Destination $backup

$excel = New-Object -ComObject Excel.Application
try {
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Open($manager, 0, $false)

    $book.Worksheets.Item("Matches").Range("I2:K25").ClearContents()
    $book.Worksheets.Item("Knockout").Range("G2:K8").ClearContents()

    $goals = $book.Worksheets.Item("Goals")
    $goals.Range("E2:E289").ClearContents()
    $goals.Range("D2:D289").FormulaR1C1 = '=IF(RC[-3]="","","Goal")'

    $knockoutGoals = $book.Worksheets.Item("KnockoutGoals")
    $knockoutGoals.Range("E2:E85").ClearContents()
    $knockoutGoals.Range("D2:D85").FormulaR1C1 = '=IF(RC[-3]="","","Goal")'

    $excel.CalculateFullRebuild()
    $excel.CalculateUntilAsyncQueriesDone()
    $book.Save()
    $book.Close($true)

    Write-Host ""
    Write-Host "Reset completato correttamente." -ForegroundColor Green
    Write-Host "Backup creato: $backup"
}
finally {
    if ($book) {
        [Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null
    }
    $excel.Quit()
    [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
