# vybe-test: powershell/json_custom_object_roundtrip/custom_object_with_timespan_property
$orig = [pscustomobject]@{ Duration = [timespan]::FromMinutes(15) }
$json = $orig | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.Duration.TotalMinutes -ne 15 -and -not $recovered.Duration.ToString().Contains("15")) {
    Write-Host "FAIL: TimeSpan property roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
