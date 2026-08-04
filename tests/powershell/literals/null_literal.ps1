# vybe-test: powershell/literals/null_literal
$value = $null
if ($value -ne $null) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
