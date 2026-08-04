# vybe-test: powershell/type_casting/string_to_datetime
$value = [datetime]'2026-08-04'
if ($value.Year -ne 2026) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
