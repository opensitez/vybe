# vybe-test: powershell/interpolation/expression_interpolation
$text = "1 + 1 = $($([int]1 + 1))"
if ($text -ne '1 + 1 = 2') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
