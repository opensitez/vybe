# vybe-test: powershell/select_string_cmdlet/select_string_quiet_returns_boolean_true
# -Quiet outputs boolean $true when a match is detected
$found = "critical system error detected" | Select-String -Pattern "error" -Quiet

if ($found -ne $true) {
    Write-Host "FAIL: Select-String -Quiet expected `$true, got: $found"
    exit 1
}

Write-Host "PASS"
exit 0
