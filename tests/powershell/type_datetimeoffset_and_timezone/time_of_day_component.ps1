# vybe-test: powershell/type_datetimeoffset_and_timezone/time_of_day_component
$dto = [datetimeoffset]::Parse("2026-08-26T08:15:30Z")
$tod = $dto.TimeOfDay
if ($tod.Hours -ne 8 -or $tod.Minutes -ne 15 -or $tod.Seconds -ne 30) {
    Write-Host "FAIL: TimeOfDay component extraction failed"
    exit 1
}
Write-Host "PASS"
exit 0
