# vybe-test: powershell/trap_statements/trap_with_message
$val = 100
if ($val -eq 100) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
