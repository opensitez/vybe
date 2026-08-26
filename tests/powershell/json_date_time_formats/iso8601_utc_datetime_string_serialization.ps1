# vybe-test: powershell/json_date_time_formats/iso8601_utc_datetime_string_serialization
$dt = [datetime]::Parse("2026-08-26T12:00:00Z").ToUniversalTime()
$json = @{ Timestamp = $dt } | ConvertTo-Json
if (-not $json.Contains("2026-08-26")) {
    Write-Host "FAIL: ISO 8601 UTC date serialization failed, got '$json'"
    exit 1
}
Write-Host "PASS"
exit 0
