# vybe-test: powershell/interpolation/multiple_variables
$first = 'A'
$second = 'B'
$text = "$first and $second"
if ($text -ne 'A and B') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
