# vybe-test: powershell/type_datetimeoffset_and_timezone/to_offset_conversion
$dto = [datetimeoffset]::Parse("2026-08-26T12:00:00+00:00")
$ny = $dto.ToOffset([timespan]::FromHours(-5))
if ($ny.Hour -ne 7 -or $ny.Offset.TotalHours -ne -5.0) {
    Write-Host "FAIL: NY offset hour expected 7, got $($ny.Hour)"
    exit 1
}
Write-Host "PASS"
exit 0
