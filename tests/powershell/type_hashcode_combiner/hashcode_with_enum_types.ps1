# vybe-test: powershell/type_hashcode_combiner/hashcode_with_enum_types
$d1 = [System.DayOfWeek]::Monday
$d2 = [System.DayOfWeek]::Monday
$h1 = [System.HashCode]::Combine($d1)
$h2 = [System.HashCode]::Combine($d2)
if ($h1 -ne $h2) { Write-Host "FAIL: HashCode with Enum failed"; exit 1 }
Write-Host "PASS"; exit 0
