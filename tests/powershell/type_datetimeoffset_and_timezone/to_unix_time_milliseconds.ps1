# vybe-test: powershell/type_datetimeoffset_and_timezone/to_unix_time_milliseconds
$dto = [datetimeoffset]::Parse("1970-01-01T00:00:01.250Z")
$ms = $dto.ToUnixTimeMilliseconds()
if ($ms -ne 1250) {
    Write-Host "FAIL: expected 1250 ms, got $ms"
    exit 1
}
Write-Host "PASS"
exit 0
