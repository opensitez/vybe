# vybe-test: powershell/error_streams/error_preference_default
if ($ErrorActionPreference -ne 'Continue') {
    Write-Host "FAIL: expected default Continue"
    exit 1
}
Write-Host 'PASS'
exit 0
