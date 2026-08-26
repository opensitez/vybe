# vybe-test: powershell/json_custom_object_roundtrip/custom_object_with_boolean_truthiness_roundtrip
$orig = [pscustomobject]@{ T = $true; F = $false }
$json = $orig | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if (-not $recovered.T -or $recovered.F) {
    Write-Host "FAIL: Boolean truthiness roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
