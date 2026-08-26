# vybe-test: powershell/type_datetimeoffset_and_timezone/to_unix_time_seconds
$dto = [datetimeoffset]::Parse("1970-01-01T01:00:00+01:00")
$unix = $dto.ToUnixTimeSeconds()
if ($unix -ne 0) {
    Write-Host "FAIL: expected 0 unix timestamp, got $unix"
    exit 1
}
Write-Host "PASS"
exit 0
