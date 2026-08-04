# vybe-test: powershell/literals/boolean_literals
$trueValue = $true
$falseValue = $false
if ($trueValue -ne $true -or $falseValue -ne $false) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
