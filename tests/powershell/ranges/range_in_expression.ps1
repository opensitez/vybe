# vybe-test: powershell/ranges/range_in_expression
$result = 0 + (1..3)
if (($result -join ',') -ne '1,2,3') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
