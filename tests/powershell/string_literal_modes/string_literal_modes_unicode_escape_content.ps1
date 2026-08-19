# vybe-test: powershell/string_literal_modes/unicode_escape_content
$pi = "\u{03A0}"
if ($pi -ne 'Π') {
    Write-Host "FAIL: expected Greek Pi, got '$pi'"
    exit 1
}

Write-Host 'PASS'
exit 0
