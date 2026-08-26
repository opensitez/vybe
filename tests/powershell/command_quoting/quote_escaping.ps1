# vybe-test: powershell/command_quoting/quote_escaping
$val = 100
if ($val -eq 100) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
