class CheckList {
    [String] $person;
    [String] $date;
    [String] $system;

    [Boolean] $hasNo1;
    [Boolean] $hasNo2;
    [Boolean] $hasNo3;
    [Boolean] $hasNo4;
    [Boolean] $hasNo5;
    [Boolean] $hasNo6;
    [Boolean] $hasNo7;
    [Boolean] $hasNo8;
    [Boolean] $hasNo9;
    [Boolean] $hasNo10;
    [Boolean] $hasNo11;
    [Boolean] $hasNo12;
    [Boolean] $hasNo13;
    [Boolean] $hasNo14;
    [Boolean] $hasNo15;
    [Boolean] $hasNo16;
    [Boolean] $hasNo17;
    [Boolean] $hasNo18;
    [Boolean] $hasNo19;
    [Boolean] $hasNo20;
    [Boolean] $hasNo21;
    [Boolean] $hasNo22;
    [Boolean] $hasNo23;
    [Boolean] $hasNo24;

    CheckList(
        [String] $person,
        [String] $date,
        [String] $system,
        [Boolean] $hasNo1,
        [Boolean] $hasNo2,
        [Boolean] $hasNo3,
        [Boolean] $hasNo4,
        [Boolean] $hasNo5,
        [Boolean] $hasNo6,
        [Boolean] $hasNo7,
        [Boolean] $hasNo8,
        [Boolean] $hasNo9,
        [Boolean] $hasNo10,
        [Boolean] $hasNo11,
        [Boolean] $hasNo12,
        [Boolean] $hasNo13,
        [Boolean] $hasNo14,
        [Boolean] $hasNo15,
        [Boolean] $hasNo16,
        [Boolean] $hasNo17,
        [Boolean] $hasNo18,
        [Boolean] $hasNo19,
        [Boolean] $hasNo20,
        [Boolean] $hasNo21,
        [Boolean] $hasNo22,
        [Boolean] $hasNo23,
        [Boolean] $hasNo24
    ) {
        $this.person = $person
        $this.date = $date
        $this.system = $system
        $this.hasNo1 = $hasNo1
        $this.hasNo2 = $hasNo2
        $this.hasNo3 = $hasNo3
        $this.hasNo4 = $hasNo4
        $this.hasNo5 = $hasNo5
        $this.hasNo6 = $hasNo6
        $this.hasNo7 = $hasNo7
        $this.hasNo8 = $hasNo8
        $this.hasNo9 = $hasNo9
        $this.hasNo10 = $hasNo10
        $this.hasNo11 = $hasNo11
        $this.hasNo12 = $hasNo12
        $this.hasNo13 = $hasNo13
        $this.hasNo14 = $hasNo14
        $this.hasNo15 = $hasNo15
        $this.hasNo16 = $hasNo16
        $this.hasNo17 = $hasNo17
        $this.hasNo18 = $hasNo18
        $this.hasNo19 = $hasNo19
        $this.hasNo20 = $hasNo20
        $this.hasNo21 = $hasNo21
        $this.hasNo22 = $hasNo22
        $this.hasNo23 = $hasNo23
        $this.hasNo24 = $hasNo24
    }

    [void] updateHasNo (
        $hasNo1, $hasNo2, $hasNo3, $hasNo4, $hasNo5, $hasNo6, $hasNo7, $hasNo8, $hasNo9, $hasNo10,
        $hasNo11, $hasNo12, $hasNo13, $hasNo14, $hasNo15, $hasNo16, $hasNo17, $hasNo18, $hasNo19, $hasNo20,
        $hasNo21, $hasNo22, $hasNo23, $hasNo24) {

        $this.hasNo1 = $hasNo1 -or $this.hasNo1
        $this.hasNo2 = $hasNo2 -or $this.hasNo2
        $this.hasNo3 = $hasNo3 -or $this.hasNo3
        $this.hasNo4 = $hasNo4 -or $this.hasNo4
        $this.hasNo5 = $hasNo5 -or $this.hasNo5
        $this.hasNo6 = $hasNo6 -or $this.hasNo6
        $this.hasNo7 = $hasNo7 -or $this.hasNo7
        $this.hasNo8 = $hasNo8 -or $this.hasNo8
        $this.hasNo9 = $hasNo9 -or $this.hasNo9
        $this.hasNo10 = $hasNo10 -or $this.hasNo10
        $this.hasNo11 = $hasNo11 -or $this.hasNo11
        $this.hasNo12 = $hasNo12 -or $this.hasNo12
        $this.hasNo13 = $hasNo13 -or $this.hasNo13
        $this.hasNo14 = $hasNo14 -or $this.hasNo14
        $this.hasNo15 = $hasNo15 -or $this.hasNo15
        $this.hasNo16 = $hasNo16 -or $this.hasNo16
        $this.hasNo17 = $hasNo17 -or $this.hasNo17
        $this.hasNo18 = $hasNo18 -or $this.hasNo18
        $this.hasNo19 = $hasNo19 -or $this.hasNo19
        $this.hasNo20 = $hasNo20 -or $this.hasNo20
        $this.hasNo21 = $hasNo21 -or $this.hasNo21
        $this.hasNo22 = $hasNo22 -or $this.hasNo22
        $this.hasNo23 = $hasNo23 -or $this.hasNo23
        $this.hasNo24 = $hasNo24 -or $this.hasNo24
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

            $hasNo1 = ![string]::IsNullOrWhiteSpace($sheet.Cells.Item($row, 4).Text)
            $hasNo2 = ![string]::IsNullOrWhiteSpace($sheet.Cells.Item($row, 5).Text)
            $hasNo3 = ![string]::IsNullOrWhiteSpace($sheet.Cells.Item($row, 6).Text)
            $hasNo4 = ![string]::IsNullOrWhiteSpace($sheet.Cells.Item($row, 7).Text)
            $hasNo5 = ![string]::IsNullOrWhiteSpace($sheet.Cells.Item($row, 8).Text)
            $hasNo6 = ![string]::IsNullOrWhiteSpace($sheet.Cells.Item($row, 9).Text)
            $hasNo7 = ![string]::IsNullOrWhiteSpace($sheet.Cells.Item($row, 10).Text)
            $hasNo8 = ![string]::IsNullOrWhiteSpace($sheet.Cells.Item($row, 11).Text)
            $hasNo9 = ![string]::IsNullOrWhiteSpace($sheet.Cells.Item($row, 12).Text)
            $hasNo10 = ![string]::IsNullOrWhiteSpace($sheet.Cells.Item($row, 13).Text)
            $hasNo11 = ![string]::IsNullOrWhiteSpace($sheet.Cells.Item($row, 14).Text)
            $hasNo12 = ![string]::IsNullOrWhiteSpace($sheet.Cells.Item($row, 15).Text)
            $hasNo13 = ![string]::IsNullOrWhiteSpace($sheet.Cells.Item($row, 16).Text)
            $hasNo14 = ![string]::IsNullOrWhiteSpace($sheet.Cells.Item($row, 17).Text)
            $hasNo15 = ![string]::IsNullOrWhiteSpace($sheet.Cells.Item($row, 18).Text)
            $hasNo16 = ![string]::IsNullOrWhiteSpace($sheet.Cells.Item($row, 19).Text)
            $hasNo17 = ![string]::IsNullOrWhiteSpace($sheet.Cells.Item($row, 20).Text)
            $hasNo18 = ![string]::IsNullOrWhiteSpace($sheet.Cells.Item($row, 21).Text)
            $hasNo19 = ![string]::IsNullOrWhiteSpace($sheet.Cells.Item($row, 22).Text)
            $hasNo20 = ![string]::IsNullOrWhiteSpace($sheet.Cells.Item($row, 23).Text)
            $hasNo21 = ![string]::IsNullOrWhiteSpace($sheet.Cells.Item($row, 24).Text)
            $hasNo22 = ![string]::IsNullOrWhiteSpace($sheet.Cells.Item($row, 25).Text)
            $hasNo23 = ![string]::IsNullOrWhiteSpace($sheet.Cells.Item($row, 26).Text)
            $hasNo24 = ![string]::IsNullOrWhiteSpace($sheet.Cells.Item($row, 27).Text)


            if (![string]::IsNullOrWhiteSpace($dateValue) -and ![string]::IsNullOrWhiteSpace($personValue)) {
                $key = "$personValue - $dateValue - $systemValue"
                if ($mapData.ContainsKey($key)) {
                    $mapData[$key].updateHasNo(
                            $hasNo1, $hasNo2, $hasNo3, $hasNo4, $hasNo5,
                            $hasNo6, $hasNo7, $hasNo8, $hasNo9, $hasNo10,
                            $hasNo11, $hasNo12, $hasNo13, $hasNo14, $hasNo15,
                            $hasNo16, $hasNo17, $hasNo18, $hasNo19, $hasNo20,
                            $hasNo21, $hasNo22, $hasNo23, $hasNo24)
                    
                } else {
                    $checkList = [CheckList]::new(
                            $personValue, $dateValue, $systemValue,
                            $hasNo1, $hasNo2, $hasNo3, $hasNo4, $hasNo5,
                            $hasNo6, $hasNo7, $hasNo8, $hasNo9, $hasNo10,
                            $hasNo11, $hasNo12, $hasNo13, $hasNo14, $hasNo15,
                            $hasNo16, $hasNo17, $hasNo18, $hasNo19, $hasNo20,
                            $hasNo21, $hasNo22, $hasNo23, $hasNo24)
                    $mapData[$key] = $checkList
                }
            }
        }

        $mapData.Values | ForEach-Object {
            Write-Host "Person: $($_.person)"
            Write-Host "Date: $($_.date)"
            Write-Host "System: $($_.system)"
            Write-Host "Has No1: $($_.hasNo1)"
            Write-Host "Has No2: $($_.hasNo2)"
            Write-Host "Has No3: $($_.hasNo3)"
            Write-Host "Has No4: $($_.hasNo4)"
            Write-Host "Has No5: $($_.hasNo5)"
            Write-Host "Has No6: $($_.hasNo6)"
            Write-Host "Has No7: $($_.hasNo7)"
            Write-Host "Has No8: $($_.hasNo8)"
            Write-Host "Has No9: $($_.hasNo9)"
            Write-Host "Has No10: $($_.hasNo10)"
            Write-Host "Has No11: $($_.hasNo11)"
            Write-Host "Has No12: $($_.hasNo12)"
            Write-Host "Has No13: $($_.hasNo13)"
            Write-Host "Has No14: $($_.hasNo14)"
            Write-Host "Has No15: $($_.hasNo15)"
            Write-Host "Has No16: $($_.hasNo16)"
            Write-Host "Has No17: $($_.hasNo17)"
            Write-Host "Has No18: $($_.hasNo18)"
            Write-Host "Has No19: $($_.hasNo19)"
            Write-Host "Has No20: $($_.hasNo20)"
            Write-Host "Has No21: $($_.hasNo21)"
            Write-Host "Has No22: $($_.hasNo22)"
            Write-Host "Has No23: $($_.hasNo23)"
            Write-Host "Has No24: $($_.hasNo24)"
            Write-Host "-------------------------------"
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