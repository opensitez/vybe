# vybe-test: powershell/type_datetimeoffset_and_timezone/subtraction_difference_timespan
$dto1 = [datetimeoffset]::Parse("2026-08-26T15:00:00Z")
$dto2 = [datetimeoffset]::Parse("2026-08-26T10:00:00Z")
$diff = $dto1 - $dto2
if ($diff.TotalHours -ne 5.0) {
    Write-Host "FAIL: expected 5 hours difference, got $($diff.TotalHours)"
    exit 1
}
Write-Host "PASS"
exit 0
