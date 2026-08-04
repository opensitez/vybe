# vybe-test: powershell/ranges/range_as_argument
$result = (1..3) -join ','
if ($result -ne '1,2,3') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
