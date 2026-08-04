# vybe-test: powershell/ternary/nested_ternary
$result = (1 -eq 2) ? 'a' : ((2 -eq 2) ? 'b' : 'c')
if ($result -ne 'b') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
