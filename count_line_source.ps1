$link = "E:\power_shell_tutor\daily-report-f-academy-form\src"
$items = Get-ChildItem -Path $link -Recurse -File
$arr_result = @()

for ($i = 0; $i -lt $items.Length; $i++) { 
    $item = $items[$i]
    $lineCount = 0

    try { 
        $lineCount = (Get-Content $item.FullName -Encoding UTF8).Count 
    } catch { 
        $lineCount = "not readable" 
    }

    $item_result = [PSCustomObject]@{
        No         = $i + 1
        File       = $item.Name
        Directory  = $item.DirectoryName
        Extension  = $item.Extension
        LineCount  = $lineCount
    }

    $arr_result += $item_result
}

$outputPath = "$link\output.csv"
$arr_result | Export-Csv -Path $outputPath -NoTypeInformation

Write-Output "Write to $outputPath"
