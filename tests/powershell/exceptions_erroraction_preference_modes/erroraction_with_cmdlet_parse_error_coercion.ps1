# vybe-test: powershell/exceptions_erroraction_preference_modes/erroraction_with_cmdlet_parse_error_coercion
$caught = $false
try {
    Get-Item "NonExistentPath_12345" -ErrorAction Stop
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Get-Item -ErrorAction Stop failed"
    exit 1
}
Write-Host "PASS"
exit 0
