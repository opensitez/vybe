# vybe-test: powershell/json_date_time_formats/datetimeoffset_json_serialization
$dto = [datetimeoffset]::Parse("2026-08-26T14:30:00+02:00")
$json = @{ Dto = $dto } | ConvertTo-Json
if (-not $json.Contains("2026-08-26")) {
    Write-Host "FAIL: DateTimeOffset JSON serialization failed"
    exit 1
}
Write-Host "PASS"
exit 0
