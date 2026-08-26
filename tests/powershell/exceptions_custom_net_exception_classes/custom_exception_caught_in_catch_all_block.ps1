# vybe-test: powershell/exceptions_custom_net_exception_classes/custom_exception_caught_in_catch_all_block
class SecretEx : System.Exception {}
$caught = $false
try {
    throw [SecretEx]::new()
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Custom exception should be caught by catch-all block"
    exit 1
}
Write-Host "PASS"
exit 0
