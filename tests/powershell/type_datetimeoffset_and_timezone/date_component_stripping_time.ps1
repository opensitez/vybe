# vybe-test: powershell/type_datetimeoffset_and_timezone/date_component_stripping_time
$dto = [datetimeoffset]::Parse("2026-08-26T18:45:30+03:00")
$dateOnly = $dto.Date
if ($dateOnly.Hour -ne 0 -or $dateOnly.Minute -ne 0 -or $dateOnly.Day -ne 26) {
    Write-Host "FAIL: Date property should zero out time components"
    exit 1
}
Write-Host "PASS"
exit 0
