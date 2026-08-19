# vybe-test: powershell/token_and_grammar_system/punctuation_array_literal
$items = @(1, 2, 3)
if ($items.Count -ne 3 -or $items[1] -ne 2) {
    Write-Host "FAIL: array literal parse/access failed, count=$($items.Count) second=$($items[1])"
    exit 1
}

Write-Host 'PASS'
exit 0
