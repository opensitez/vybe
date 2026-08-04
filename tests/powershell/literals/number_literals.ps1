# vybe-test: powershell/literals/number_literals
$int = 42
$hex = 0x2A
if ($int -ne 42 -or $hex -ne 42) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
