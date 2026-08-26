# vybe-test: powershell/should_process/should_process_should_continue
$val = 100
if ($val -eq 100) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
