# vybe-test: powershell/type_converters/type_converter_enum_string
$day = [System.DayOfWeek]"Monday"
if ($day -ne [System.DayOfWeek]::Monday) {
    Write-Host "FAIL: string to enum [DayOfWeek] conversion failed"
    exit 1
}
Write-Host "PASS"
exit 0
