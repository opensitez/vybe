# vybe-test: powershell/type_datetimeoffset_and_timezone/equality_same_point_in_time
$dto1 = [datetimeoffset]::Parse("2026-01-01T00:00:00+00:00")
$dto2 = [datetimeoffset]::Parse("2025-12-31T19:00:00-05:00")
if ($dto1 -ne $dto2) {
    Write-Host "FAIL: instants in time should be equal"
    exit 1
}
Write-Host "PASS"
exit 0
