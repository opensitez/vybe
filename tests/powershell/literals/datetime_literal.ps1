# vybe-test: powershell/literals/datetime_literal
$date = [datetime]'2026-01-01'
if ($date.Year -ne 2026) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
