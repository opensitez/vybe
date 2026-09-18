# vybe-test: powershell/join_string_cmdlet/join_string_single_quote_each_element
# -SingleQuote encloses each individual element in single quotes before joining
$res = @('one', 'two') | Join-String -Separator ', ' -SingleQuote

$expected = "'one', 'two'"

if ($res -ne $expected) {
    Write-Host "FAIL: expected $expected, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
