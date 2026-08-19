# vybe-test: powershell/string_substitution_rules/substitute_calculation
$x = 4
$y = 5
$result = "$( $x * $y )"
if ($result -ne '20') {
    Write-Host "FAIL: calculation substitution expected 20, got '$result'"
    exit 1
}

Write-Host 'PASS'
exit 0
