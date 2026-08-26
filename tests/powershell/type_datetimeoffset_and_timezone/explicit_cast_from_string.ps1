# vybe-test: powershell/type_datetimeoffset_and_timezone/explicit_cast_from_string
$dto = [datetimeoffset]"2026-12-31T23:59:59Z"
if ($dto.Year -ne 2026 -or $dto.Month -ne 12 -or $dto.Day -ne 31) {
    Write-Host "FAIL: type accelerator cast to DateTimeOffset failed"
    exit 1
}
Write-Host "PASS"
exit 0
