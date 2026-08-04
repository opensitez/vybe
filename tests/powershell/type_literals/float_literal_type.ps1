# vybe-test: powershell/type_literals/float_literal_type
$value = [float]1.5
if ($value -ne 1.5) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
