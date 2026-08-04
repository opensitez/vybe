# vybe-test: powershell/ternary/false_condition
$result = (1 -eq 2) ? 'yes' : 'no'
if ($result -ne 'no') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
