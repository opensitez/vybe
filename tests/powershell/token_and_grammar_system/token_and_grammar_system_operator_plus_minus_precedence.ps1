# vybe-test: powershell/token_and_grammar_system/operator_plus_minus_precedence
$result = 8 + 4 - 2
if ($result -ne 10) {
    Write-Host "FAIL: plus/minus precedence wrong, got $result"
    exit 1
}

$result2 = 8 - 4 + 2
if ($result2 -ne 6) {
    Write-Host "FAIL: left-to-right arithmetic precedence wrong, got $result2"
    exit 1
}

Write-Host 'PASS'
exit 0
