# vybe-test: powershell/enums/enum_get_values
enum Color { Red; Green; Blue; Yellow }
$values = [Enum]::GetValues([Color])
if ($values.Count -ne 4) { Write-Host "FAIL: count $($values.Count)"; exit 1 }
$names = [Enum]::GetNames([Color])
if ("Green" -notin $names) { Write-Host "FAIL: Green not in names"; exit 1 }
Write-Host "PASS"
exit 0
