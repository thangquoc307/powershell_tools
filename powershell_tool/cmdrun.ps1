function Run-CmdWithCommands {
    param (
        [string[]]$Commands
    )
    
    # Tạo file batch tạm để chạy các lệnh
    $tempBatchFile = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.bat'

    # Ghi các lệnh vào file batch
    $Commands | Out-File -FilePath $tempBatchFile -Encoding ASCII

    # Chạy file batch bằng cmd.exe
    Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$tempBatchFile`"" -NoNewWindow -Wait

    # Xóa file batch tạm
    Remove-Item -Path $tempBatchFile -Force
}

# Sử dụng hàm
Run-CmdWithCommands -Commands @(
    "echo Xin!",
    "echo alo."
)
