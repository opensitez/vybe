# vybe-test: powershell/type_literals/enum_literal_type
$value = [System.DayOfWeek]::Friday
if ($value -ne [System.DayOfWeek]::Friday) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
