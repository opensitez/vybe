# vybe-test: powershell/join_string_cmdlet/join_string_combining_quotes_prefix_suffix
# Combining -SingleQuote, -Separator, -OutputPrefix, and -OutputSuffix produces an enclosed quoted list
$res = @('first', 'second') | Join-String -Separator ', ' -SingleQuote -OutputPrefix '(' -OutputSuffix ')'

$expected = "('first', 'second')"

if ($res -ne $expected) {
    Write-Host "FAIL: expected $expected, got '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
