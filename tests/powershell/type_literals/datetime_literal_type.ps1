# vybe-test: powershell/type_literals/datetime_literal_type
$value = [datetime]'2026-08-04'
if ($value.Year -ne 2026) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
