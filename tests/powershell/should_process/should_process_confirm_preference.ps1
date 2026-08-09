# vybe-test: powershell/should_process/should_process_confirm_preference
if ($ConfirmPreference -ne $null) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL: \$ConfirmPreference default variable missing"
exit 1
