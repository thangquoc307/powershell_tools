function changeDate {
    param (
        [String] $filePath,
        [String] $date
    )
    process {
        $changingDate = Get-Date $date
        $(Get-Item $filePath).CreationTime = $changingDate
        $(Get-Item $filePath).LastWriteTime = $changingDate

        Write-Output "Updated CreationTime and LastWriteTime for $filePath"
    }
}

changeDate -filePath "E:\power_shell_tutor\CheckListAnalyze.xlsx" -date "2000-01-01 08:00:00"