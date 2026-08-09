# vybe-test: powershell/type_converters/type_converter_enum_int
$day = [System.DayOfWeek]1
if ($day -ne [System.DayOfWeek]::Monday) {
    Write-Host "FAIL: int to enum [DayOfWeek] conversion expected Monday, got $day"
    exit 1
}
Write-Host "PASS"
exit 0
