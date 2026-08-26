# vybe-test: powershell/parameters_validate_range/validaterange_error_message_contains_bounds
function Set-Day {
    param([ValidateRange(1, 31)][int]$Day)
    return $Day
}
$msg = ""
try {
    $x = Set-Day -Day 35
} catch {
    $msg = $_.Exception.Message
}
if (-not ($msg.Contains("1") -and $msg.Contains("31") -and $msg.Contains("35"))) {
    Write-Host "FAIL: Error message should contain min, max, and actual value, got '$msg'"
    exit 1
}
Write-Host "PASS"
exit 0
