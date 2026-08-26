# vybe-test: powershell/type_datetimeoffset_and_timezone/comparison_same_instant_different_offsets
$dto1 = [datetimeoffset]::Parse("2026-08-26T14:00:00+02:00")
$dto2 = [datetimeoffset]::Parse("2026-08-26T12:00:00Z")
if ($dto1 -ne $dto2) {
    Write-Host "FAIL: 14:00+2 should equal 12:00Z in universal comparison"
    exit 1
}
Write-Host "PASS"
exit 0
