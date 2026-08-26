# vybe-test: powershell/type_datetimeoffset_and_timezone/hour_minute_second_offset_extraction
$dto = [datetimeoffset]::new(2026, 5, 10, 18, 45, 12, [timespan]::FromHours(3))
if ($dto.Hour -ne 18 -or $dto.Minute -ne 45 -or $dto.Second -ne 12 -or $dto.Offset.Hours -ne 3) {
    Write-Host "FAIL: Time component extraction failed"
    exit 1
}
Write-Host "PASS"
exit 0
