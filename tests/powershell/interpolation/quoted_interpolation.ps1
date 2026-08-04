# vybe-test: powershell/interpolation/quoted_interpolation
$value = 'x'
$text = "'$value'"
if ($text -ne "'x'") {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
