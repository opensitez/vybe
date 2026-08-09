# vybe-test: powershell/variable_drives/variable_drive_global_prefix
$global:vybeGlobalTestVar = 999
if ($global:vybeGlobalTestVar -ne 999) {
    Write-Host "FAIL: \$global: prefix read failed"
    exit 1
}
Write-Host "PASS"
exit 0
