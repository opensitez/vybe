# vybe-test: powershell/ranges/decreasing_range
$range = 5..1
if ($range.Length -ne 5 -or $range[0] -ne 5) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
