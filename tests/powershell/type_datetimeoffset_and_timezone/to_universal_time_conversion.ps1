# vybe-test: powershell/type_datetimeoffset_and_timezone/to_universal_time_conversion
$dto = [datetimeoffset]::Parse("2026-08-26T14:30:00+02:00")
$utc = $dto.ToUniversalTime()
if ($utc.Hour -ne 12 -or $utc.Offset.TotalHours -ne 0.0) {
    Write-Host "FAIL: UTC hour expected 12, got $($utc.Hour)"
    exit 1
}
Write-Host "PASS"
exit 0
