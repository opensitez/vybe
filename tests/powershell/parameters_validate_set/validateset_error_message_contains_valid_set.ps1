# vybe-test: powershell/parameters_validate_set/validateset_error_message_contains_valid_set
function Set-Role {
    param([ValidateSet("User", "Admin")][string]$Role)
    return $Role
}
$msg = ""
try {
    $x = Set-Role -Role "Guest"
} catch {
    $msg = $_.Exception.Message
}
if (-not ($msg.Contains("User") -and $msg.Contains("Admin") -and $msg.Contains("Guest"))) {
    Write-Host "FAIL: Error message should list allowed values and provided value, got '$msg'"
    exit 1
}
Write-Host "PASS"
exit 0
