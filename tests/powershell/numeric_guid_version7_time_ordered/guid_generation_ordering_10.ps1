# vybe-test: powershell/numeric_guid_version7_time_ordered/guid_generation_ordering_10
$g1 = [guid]::NewGuid()
$g2 = [guid]::NewGuid()
if ($g1 -eq $g2 -or $g1.ToString().Length -ne 36) { Write-Host "FAIL: Guid uniqueness failed"; exit 1 }
Write-Host "PASS"; exit 0
