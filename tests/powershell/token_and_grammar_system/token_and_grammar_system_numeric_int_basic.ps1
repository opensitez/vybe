# vybe-test: powershell/token_and_grammar_system/numeric_int_basic
$value = 12345
if ($value -ne 12345) {
    Write-Host "FAIL: integer literal parsed incorrectly ($value)"
    exit 1
}

if ($value.GetType().Name -ne 'Int32') {
    Write-Host "FAIL: integer literal type is not Int32"
    exit 1
}

Write-Host 'PASS'
exit 0
