# vybe-test: powershell/type_literals/object_literal_type
$value = [object]'hello'
if ($value -ne 'hello') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
