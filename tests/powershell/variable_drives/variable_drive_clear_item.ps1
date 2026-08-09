# vybe-test: powershell/variable_drives/variable_drive_clear_item
$toClear = "Something"
Clear-Item "variable:toClear"
if ($toClear -ne $null) {
    Write-Host "FAIL: Clear-Item variable:toClear expected null, got $toClear"
    exit 1
}
Write-Host "PASS"
exit 0
