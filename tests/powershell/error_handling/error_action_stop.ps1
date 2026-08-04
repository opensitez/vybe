# vybe-test: powershell/error_handling/error_action_stop
$ErrorActionPreference = "Stop"
$caught = $false
try {
    throw "test"
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: expected error to be caught"
    exit 1
}
Write-Host "PASS"
exit 0
