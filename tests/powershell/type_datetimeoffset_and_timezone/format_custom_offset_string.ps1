# vybe-test: powershell/type_datetimeoffset_and_timezone/format_custom_offset_string
$dto = [datetimeoffset]::new(2026, 4, 15, 9, 30, 0, [timespan]::FromHours(-4))
$str = $dto.ToString("yyyy-MM-dd HH:mm zzz")
if ($str -ne "2026-04-15 09:30 -04:00") {
    Write-Host "FAIL: expected '2026-04-15 09:30 -04:00', got '$str'"
    exit 1
}
Write-Host "PASS"
exit 0
