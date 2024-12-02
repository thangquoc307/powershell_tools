$winmergepath = "C:\Program Files\WinMerge\WinMergeU.exe"
$file1 = "E:\power_shell_tutor\testwinmerge\New folder\test_readline_ex_cmt.pc"
$file2 = "E:\power_shell_tutor\testwinmerge\New folder\test_readline_ex_cmtcopy.pc"

$reportPath = "E:\power_shell_tutor\testwinmerge\ComparisonReport.html"



& $winmergepath $file1 $file2 /r /e /u /wl /wr /x /o $reportPath
