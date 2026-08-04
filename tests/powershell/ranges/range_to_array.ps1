# vybe-test: powershell/ranges/range_to_array
$array = 1..3
if ($array[1] -ne 2) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
