# vybe-test: powershell/ternary/ternary_with_expressions
$result = (1 + 1 -eq 2) ? 'yes' : 'no'
if ($result -ne 'yes') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
