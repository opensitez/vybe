# vybe-test: powershell/json_date_time_formats/unix_timestamp_seconds_integer_format
$epochSeconds = 1787660400 # 2026-08-26 approx
$json = @{ Epoch = $epochSeconds } | ConvertTo-Json
$obj = $json | ConvertFrom-Json
$dt = [datetimeoffset]::FromUnixTimeSeconds($obj.Epoch).UtcDateTime
if ($dt.Year -ne 2026) {
    Write-Host "FAIL: Unix timestamp seconds roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
