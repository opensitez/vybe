# vybe-test: powershell/string_literal_modes/formatting_with_braces
$result = "{0,6}" -f 'ok'
if ($result -ne '    ok') {
    Write-Host "FAIL: format alignment failed: '$result'"
    exit 1
}

Write-Host 'PASS'
exit 0
