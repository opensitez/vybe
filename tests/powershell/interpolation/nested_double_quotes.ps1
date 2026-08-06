# vybe-test: powershell/interpolation/nested_double_quotes
# A `$( … )` inside a double-quoted string holds CODE, so that code may spell
# its OWN double-quoted strings. The inner quote must not end the outer string.
$text = "upper=$("lit".ToUpper())"
if ($text -ne 'upper=LIT') {
    Write-Host "FAIL: got [$text]"
    exit 1
}
Write-Host 'PASS'
exit 0
