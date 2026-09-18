# vybe-test: powershell/filesystem_item_cmdlets/get_item_on_existing_file_returns_fileinfo
# Get-Item on an existing file returns a System.IO.FileInfo object with correct Name and Exists
$tmp = [System.IO.Path]::GetTempFileName()
"content" | Set-Content $tmp

$item = Get-Item $tmp

if ($item.GetType().Name -ne "FileInfo") {
    Write-Host "FAIL: expected FileInfo, got $($item.GetType().Name)"
    Remove-Item $tmp -Force
    exit 1
}

if (-not $item.Exists) {
    Write-Host "FAIL: FileInfo.Exists is false"
    Remove-Item $tmp -Force
    exit 1
}

Remove-Item $tmp -Force
Write-Host "PASS"
exit 0
