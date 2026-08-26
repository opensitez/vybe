# vybe-test: powershell/type_datetimeoffset_and_timezone/day_of_week_and_day_of_year
$dto = [datetimeoffset]::Parse("2026-01-01T00:00:00Z")
if ($dto.DayOfYear -ne 1 -or $dto.DayOfWeek -ne [System.DayOfWeek]::Thursday) {
    Write-Host "FAIL: DayOfWeek/DayOfYear mismatch for 2026-01-01"
    exit 1
}
Write-Host "PASS"
exit 0
