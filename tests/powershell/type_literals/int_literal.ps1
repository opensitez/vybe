# vybe-test: powershell/type_literals/int_literal
$value = [int]5
if ($value -ne 5) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
