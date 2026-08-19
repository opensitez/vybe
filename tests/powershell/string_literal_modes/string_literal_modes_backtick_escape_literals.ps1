# vybe-test: powershell/string_literal_modes/backtick_escape_literals
$value = "literal `"quoted`" `$x and `` symbol"
if ($value -ne 'literal "quoted" $x and ` symbol') {
    Write-Host "FAIL: literal backtick escapes did not produce expected text: $value"
    exit 1
}

Write-Host 'PASS'
exit 0
