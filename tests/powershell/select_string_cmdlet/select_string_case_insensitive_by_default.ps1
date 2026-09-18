# vybe-test: powershell/select_string_cmdlet/select_string_case_insensitive_by_default
# Pattern matching in Select-String is case-insensitive by default
$match = "SERVER_HOSTNAME_PRIMARY" | Select-String -Pattern "server_hostname"

if ($null -eq $match) {
    Write-Host "FAIL: case-insensitive pattern matching failed to find match"
    exit 1
}

Write-Host "PASS"
exit 0
