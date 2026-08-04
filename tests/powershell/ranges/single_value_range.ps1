# vybe-test: powershell/ranges/single_value_range
$range = 5..5
if ($range.Length -ne 1 -or $range[0] -ne 5) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
