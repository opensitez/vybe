# vybe-test: powershell/json_date_time_formats/unix_timestamp_milliseconds_integer_format
$epochMs = 1787660400000
$json = @{ EpochMs = $epochMs } | ConvertTo-Json
$obj = $json | ConvertFrom-Json
$dt = [datetimeoffset]::FromUnixTimeMilliseconds($obj.EpochMs).UtcDateTime
if ($dt.Year -ne 2026) {
    Write-Host "FAIL: Unix timestamp milliseconds roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
