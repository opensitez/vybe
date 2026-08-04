# vybe-test: powershell/ranges/string_range
$range = 'a'..'c'
if ($range.Length -ne 3 -or $range[0] -ne 'a') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
