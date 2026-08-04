# vybe-test: powershell/type_literals/byte_literal_type
$value = [byte]1
if ($value -ne 1) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
