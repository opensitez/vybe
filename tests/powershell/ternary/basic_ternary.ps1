# vybe-test: powershell/ternary/basic_ternary
$result = (1 -eq 1) ? 'yes' : 'no'
if ($result -ne 'yes') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
