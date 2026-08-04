# vybe-test: powershell/interpolation/complex_expression
$text = "Sum: $($([int]2 + 3))"
if ($text -ne 'Sum: 5') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
