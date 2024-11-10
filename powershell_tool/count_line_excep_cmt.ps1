function count_line_excep_cmt {
    param (
        [String]$Url,
        [String]$SingleCmt,
        [String]$StartMultiCmt,
        [String]$EndMultiCmt
    )
    process {
        $contents = Get-Content $Url -Encoding UTF8
        $count = 0
        $onMultiesCmt = $false
        for ($i = 0; $i -lt $contents.Length; $i ++) {
            $content = $contents[$i].Trim()
            if ($content.Length -eq 0) {
                continue
            } elseif ($onMultiesCmt) {
                if ($content.Length -ge $EndMultiCmt.Length -and $content.Substring($EndMultiCmt.Length-2) -eq $EndMultiCmt) {
                    $onMultiesCmt = $false
                }
                continue
            } else {
                if ($content.Length -ge $StartMultiCmt.Length -and $content.Substring(0, $StartMultiCmt.Length) -eq $StartMultiCmt) {
                    if ($content.Substring($SingleCmt.Length-2) -ne $EndMultiCmt) {
                        $onMultiesCmt = $true
                    }
                    continue
                } elseif ($content.Length -ge $SingleCmt.Length -and $content.Substring(0, $SingleCmt.Length) -eq $SingleCmt) {
                    continue
                } else {
                    $count++
                }
            }
        }
        Write-Output "$count lines"
    }
}
count_line_excep_cmt -Url "E:\power_shell_tutor\test_readline_ex_cmt.pc" -SingleCmt "//" -StartMultiCmt "/*" -EndMultiCmt "*/"