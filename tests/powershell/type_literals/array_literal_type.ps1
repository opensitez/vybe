# vybe-test: powershell/type_literals/array_literal_type
$value = [int[]](1,2,3)
if ($value.Length -ne 3) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
