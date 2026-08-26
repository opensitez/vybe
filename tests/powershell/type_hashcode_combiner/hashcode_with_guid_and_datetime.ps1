# vybe-test: powershell/type_hashcode_combiner/hashcode_with_guid_and_datetime
$g = [guid]::NewGuid()
$dt = [datetime]::UtcNow
$h1 = [System.HashCode]::Combine($g, $dt)
$h2 = [System.HashCode]::Combine($g, $dt)
if ($h1 -ne $h2) { Write-Host "FAIL: HashCode with Guid/DateTime failed"; exit 1 }
Write-Host "PASS"; exit 0
