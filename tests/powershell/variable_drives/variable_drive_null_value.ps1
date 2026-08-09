# vybe-test: powershell/variable_drives/variable_drive_null_value
$variable:nullItem = $null
if ($nullItem -ne $null) {
    Write-Host "FAIL: \$variable:nullItem expected null"
    exit 1
}
Write-Host "PASS"
exit 0
