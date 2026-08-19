# vybe-test: powershell/token_and_grammar_system/separator_newline_expression
$x = 10
$y = 4
$z = $x + $y
if ($z -ne 14) {
    Write-Host "FAIL: newline-separated expression failed with z=$z"
    exit 1
}

Write-Host 'PASS'
exit 0
