# vybe-test: powershell/ranges/inclusive_range
$range = 1..3
if ($range.Length -ne 3 -or $range[2] -ne 3) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
