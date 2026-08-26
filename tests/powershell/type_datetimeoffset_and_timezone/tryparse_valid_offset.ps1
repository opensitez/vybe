# vybe-test: powershell/type_datetimeoffset_and_timezone/tryparse_valid_offset
$dto = [datetimeoffset]::MinValue
$success = [datetimeoffset]::TryParse("2026-08-26 10:00:00 +00:00", [ref]$dto)
if (-not $success -or $dto.Year -ne 2026) {
    Write-Host "FAIL: TryParse DateTimeOffset failed"
    exit 1
}
Write-Host "PASS"
exit 0
