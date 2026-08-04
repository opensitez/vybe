# vybe-test: powershell/type_literals/string_literal_type
$value = [string]'hello'
if ($value -ne 'hello') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
