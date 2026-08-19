# vybe-test: powershell/token_and_grammar_system/separator_semicolon_optional
$a = 2; $b = 3
$c = $a + $b
if ($c -ne 5) {
    Write-Host "FAIL: semicolon optional behavior failed with c=$c"
    exit 1
}

Write-Host 'PASS'
exit 0
