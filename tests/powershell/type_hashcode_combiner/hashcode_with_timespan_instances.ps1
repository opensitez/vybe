# vybe-test: powershell/type_hashcode_combiner/hashcode_with_timespan_instances
$ts1 = [timespan]::FromHours(5)
$ts2 = [timespan]::FromHours(5)
$h1 = [System.HashCode]::Combine($ts1)
$h2 = [System.HashCode]::Combine($ts2)
if ($h1 -ne $h2) { Write-Host "FAIL: HashCode with TimeSpan failed"; exit 1 }
Write-Host "PASS"; exit 0
