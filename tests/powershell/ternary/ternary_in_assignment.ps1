# vybe-test: powershell/ternary/ternary_in_assignment
$x = 0
$result = ($x -eq 0) ? 'zero' : 'nonzero'
if ($result -ne 'zero') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
