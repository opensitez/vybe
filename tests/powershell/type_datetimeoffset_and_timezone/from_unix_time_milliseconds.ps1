# vybe-test: powershell/type_datetimeoffset_and_timezone/from_unix_time_milliseconds
$dto = [datetimeoffset]::FromUnixTimeMilliseconds(1500)
if ($dto.Millisecond -ne 500 -or $dto.Second -ne 1) {
    Write-Host "FAIL: FromUnixTimeMilliseconds extraction failed"
    exit 1
}
Write-Host "PASS"
exit 0
