# vybe-test: powershell/exceptions_erroraction_preference_modes/erroraction_on_convertfrom_json_cmdlet
$caught = $false
try {
    ConvertFrom-Json "{ invalid json }" -ErrorAction Stop
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: ConvertFrom-Json -ErrorAction Stop failed"
    exit 1
}
Write-Host "PASS"
exit 0
