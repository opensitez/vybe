# vybe-test: powershell/token_and_grammar_system/operator_minus_unary_vs_binary
$unary = -7
$binary = 10 - 3
if ($unary -ne -7) {
    Write-Host "FAIL: unary minus wrong, got $unary"
    exit 1
}

if ($binary -ne 7) {
    Write-Host "FAIL: binary minus wrong, got $binary"
    exit 1
}

Write-Host 'PASS'
exit 0
