# vybe-test: powershell/join_string_cmdlet/join_string_double_quote_each_element
# -DoubleQuote encloses each individual element in double quotes before joining
$res = @('alpha', 'beta') | Join-String -Separator ', ' -DoubleQuote

$expected = '"alpha", "beta"'

if ($res -ne $expected) {
    Write-Host "FAIL: expected $expected, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
