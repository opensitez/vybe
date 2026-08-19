# vybe-test: powershell/token_and_grammar_system/numeric_hex_basic
$value = 0x2A
if ($value -ne 42) {
    Write-Host "FAIL: hex literal should be 42, got $value"
    exit 1
}

if ($value.GetType().Name -notin @('Int32','Int64')) {
    Write-Host "FAIL: hex literal is not integer type"
    exit 1
}

Write-Host 'PASS'
exit 0
