# vybe-test: powershell/join_string_cmdlet/join_string_whitespace_preservation
$spaced = @('  spaced  ', 'tab`there')

# Whitespace inside elements must be strictly preserved
$res = $spaced | Join-String -Separator ';' -DoubleQuote

$expected = '"  spaced  ";"tab`there"'

if ($res -ne $expected) {
    Write-Host "FAIL: whitespace was not preserved, expected '$expected', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
