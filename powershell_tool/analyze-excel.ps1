class CheckList {
    [String] $person;
    [String] $date;
    [String] $system;

    [Boolean[]] $hasNo;

    CheckList(
        [String] $person,
        [String] $date,
        [String] $system,
        [Boolean[]] $hasNo
    ) {
        $this.person = $person
        $this.date = $date
        $this.system = $system
        $this.hasNo = $hasNo
    }
}

function computingFile {
    param (
        [String] $filePath,
        [String] $sheetName,
        [int] $startRow
    )
    process {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false

        $workbook = $excel.Workbooks.Open($filePath)
        $sheet = $workbook.Sheets.Item($sheetName)

        $usedRange = $sheet.UsedRange
        $rowCount = $usedRange.Rows.Count
        $mapData = @{}

        for ($row = $startRow; $row -le $rowCount; $row++) {
            $systemValue = $sheet.Cells.Item($row, 2).Text
            $dateValue = $sheet.Cells.Item($row, 33).Text
            $personValue = $sheet.Cells.Item($row, 31).Text
            $hasNo = @()

            for ([int]$i = 4; $i -le 27; $i++) {
                $hasNo += ![string]::IsNullOrWhiteSpace($sheet.Cells.Item($row, $i).Text)
            } 

            if (![string]::IsNullOrWhiteSpace($dateValue) -and ![string]::IsNullOrWhiteSpace($personValue)) {
                $key = "$personValue - $dateValue - $systemValue"
                if ($mapData.ContainsKey($key)) {
                    $thisHasNo = $mapData[$key].hasNo
                    for ([int]$i = 0; $i -lt $thisHasNo.Length; $i++) {
                        $thisHasNo[$i] = $thisHasNo[$i] -or $hasNo[$i]
                    }
                } else {
                    $checkList = [CheckList]::new($personValue, $dateValue, $systemValue, $hasNo)
                    $mapData[$key] = $checkList
                }
            }
        }
        foreach ($obj in $mapData.Values) {
            Write-Output "$obj.system"
            $obj.hasNo
        }


        #$workbook.Save()
        $workbook.Close()
        $excel.Quit()

        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

computingFile -filePath "E:\power_shell_tutor\CheckListAnalyze.xlsx" -sheetName "Sheet1" -startRow 3