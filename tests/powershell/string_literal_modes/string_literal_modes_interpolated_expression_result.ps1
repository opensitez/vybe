# vybe-test: powershell/string_literal_modes/interpolated_expression_result
$a = 2
$b = 3
$result = "$($a + $b)"
if ($result -ne '5') {
    Write-Host "FAIL: arithmetic interpolation result mismatch: '$result'"
    exit 1
}

Write-Host 'PASS'
exit 0
