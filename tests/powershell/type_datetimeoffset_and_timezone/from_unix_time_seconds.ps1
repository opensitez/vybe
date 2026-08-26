# vybe-test: powershell/type_datetimeoffset_and_timezone/from_unix_time_seconds
$dto = [datetimeoffset]::FromUnixTimeSeconds(0)
if ($dto.UtcDateTime.Year -ne 1970 -or $dto.UtcDateTime.Month -ne 1 -or $dto.UtcDateTime.Day -ne 1) {
    Write-Host "FAIL: Unix epoch 0 should be 1970-01-01"
    exit 1
}
Write-Host "PASS"
exit 0
