# vybe-test: powershell/compound_assignment/mod_assignment
$x = 5
$x %= 2
if ($x -ne 1) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
