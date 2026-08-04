# vybe-test: powershell/strings/string_verbatim_single_quotes
$name = "Alice"
$literal = 'Hello, $name! No interpolation here.'
if ($literal -ne 'Hello, $name! No interpolation here.') {
    Write-Host "FAIL: single quotes should not interpolate"
    exit 1
}
if ($literal -eq "Hello, Alice! No interpolation here.") {
    Write-Host "FAIL: variable should NOT be expanded"
    exit 1
}
Write-Host "PASS"
exit 0
