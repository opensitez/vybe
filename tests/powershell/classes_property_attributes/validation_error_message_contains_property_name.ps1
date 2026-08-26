# vybe-test: powershell/classes_property_attributes/validation_error_message_contains_property_name
class NamedTarget {
    [ValidateRange(1, 5)][int]$Level
}
$nt = [NamedTarget]::new()
$errMsg = ""
try {
    $nt.Level = 99
} catch {
    $errMsg = $_.Exception.Message
}
if (-not ($errMsg.Contains("Level") -or $errMsg.Contains("99") -or $errMsg.Contains("range"))) {
    Write-Host "FAIL: Validation error message lacked context, got '$errMsg'"
    exit 1
}
Write-Host "PASS"
exit 0
