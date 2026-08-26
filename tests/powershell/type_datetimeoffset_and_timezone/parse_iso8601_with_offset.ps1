# vybe-test: powershell/type_datetimeoffset_and_timezone/parse_iso8601_with_offset
$str = "2026-08-26T14:30:00+02:00"
$dto = [datetimeoffset]::Parse($str)
if ($dto.Year -ne 2026 -or $dto.Month -ne 8 -or $dto.Day -ne 26 -or $dto.Hour -ne 14 -or $dto.Offset.TotalHours -ne 2.0) {
    Write-Host "FAIL: parsed DateTimeOffset properties mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
