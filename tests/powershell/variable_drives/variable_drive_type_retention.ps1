# vybe-test: powershell/variable_drives/variable_drive_type_retention
$variable:typedInt = [int]42
if (-not ($typedInt -is [int])) {
    Write-Host "FAIL: \$variable: typed int assignment lost type"
    exit 1
}
Write-Host "PASS"
exit 0
