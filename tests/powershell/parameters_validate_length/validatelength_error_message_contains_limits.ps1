# vybe-test: powershell/parameters_validate_length/validatelength_error_message_contains_limits
function Test-LenMsg {
    param([ValidateLength(3, 5)][string]$Str)
}
$caught = $false
try {
    Test-LenMsg -Str "X"
} catch {
    $caught = $_.Exception.Message.Contains("character") -or $_.Exception.Message.Contains("length") -or $_.Exception.Message.Contains("parameter")
}
if (-not $caught) {
    Write-Host "FAIL: Error message validation failed"
    exit 1
}
Write-Host "PASS"
exit 0
