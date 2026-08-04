# vybe-test: powershell/ternary/ternary_with_variables
$x = 10
$result = ($x -gt 5) ? 'big' : 'small'
if ($result -ne 'big') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
