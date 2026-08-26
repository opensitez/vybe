# vybe-test: powershell/type_datetimeoffset_and_timezone/addition_of_timespan
$dto = [datetimeoffset]::Parse("2026-08-26T10:00:00Z")
$result = $dto.AddHours(3.5)
if ($result.Hour -ne 13 -or $result.Minute -ne 30) {
    Write-Host "FAIL: expected 13:30, got $($result.Hour):$($result.Minute)"
    exit 1
}
Write-Host "PASS"
exit 0
