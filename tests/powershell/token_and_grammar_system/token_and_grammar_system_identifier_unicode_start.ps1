# vybe-test: powershell/token_and_grammar_system/identifier_unicode_start
$α = 41
$β = 1
if (($α + $β) -ne 42) {
    Write-Host "FAIL: unicode identifiers should be usable, got $($α + $β)"
    exit 1
}

Write-Host 'PASS'
exit 0
