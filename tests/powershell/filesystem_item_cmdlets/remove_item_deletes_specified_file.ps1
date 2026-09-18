# vybe-test: powershell/filesystem_item_cmdlets/remove_item_deletes_specified_file
# Remove-Item deletes an existing file so that Test-Path evaluates to false
$tmp = [System.IO.Path]::GetTempFileName()
"temporary line" | Set-Content $tmp

if (-not (Test-Path $tmp)) {
    Write-Host "FAIL: precondition failed: temp file does not exist"
    exit 1
}

Remove-Item -Path $tmp

if (Test-Path $tmp) {
    Write-Host "FAIL: file still exists after Remove-Item"
    Remove-Item $tmp -Force
    exit 1
}

Write-Host "PASS"
exit 0
