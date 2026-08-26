# vybe-test: powershell/type_datetimeoffset_and_timezone/year_month_day_extraction
$dto = [datetimeoffset]::new(2030, 11, 25, 8, 15, 30, [timespan]::Zero)
if ($dto.Year -ne 2030 -or $dto.Month -ne 11 -or $dto.Day -ne 25) {
    Write-Host "FAIL: Date component extraction failed"
    exit 1
}
Write-Host "PASS"
exit 0
