$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$filePath = "E:\power_shell_tutor\Dom.xlsx"
$sheetNameToDelete = "Sheet2"

$workbook = $excel.Workbooks.Open($filePath)
$sheet = $workbook.Sheets.Item("Sheet1")

$cell = $sheet.Range("I1")
$cell.Value2 = "3"

$excel.Calculate()

$usedRange = $sheet.UsedRange
foreach ($cell in $usedRange) {
    if ($cell.HasFormula) {
        $cell.Value2 = $cell.Value2
    }
}

foreach ($sheet in $workbook.Sheets) {
    if ($sheet.Name -eq $sheetNameToDelete) {
        $sheet.Delete()
        break
    }
}

$workbook.Save()
$workbook.Close()
$excel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()