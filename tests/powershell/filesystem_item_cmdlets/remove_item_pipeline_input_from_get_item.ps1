# vybe-test: powershell/filesystem_item_cmdlets/remove_item_pipeline_input_from_get_item
# Get-Item pipeline output can be piped directly into Remove-Item to delete the item
$tmp = [System.IO.Path]::GetTempFileName()
"piped removal target" | Set-Content $tmp

Get-Item $tmp | Remove-Item

if (Test-Path $tmp) {
    Write-Host "FAIL: file still exists after Get-Item | Remove-Item"
    Remove-Item $tmp -Force
    exit 1
}

Write-Host "PASS"
exit 0
