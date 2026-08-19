# vybe-test: powershell/token_and_grammar_system/punctuation_subscript_bounds
$list = @(10, 20, 30)
$missing = $list[99]
if ($null -ne $missing) {
    Write-Host "FAIL: out-of-range index should return null, got $missing"
    exit 1
}

Write-Host 'PASS'
exit 0
