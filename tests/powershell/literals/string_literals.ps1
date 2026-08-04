# vybe-test: powershell/literals/string_literals
$single = 'hello'
$double = "world"
if ($single -ne 'hello' -or $double -ne 'world') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
