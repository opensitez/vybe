# vybe-test: powershell/type_hashcode_combiner/hashcode_with_boolean_types
$hTrue = [System.HashCode]::Combine($true)
$hFalse = [System.HashCode]::Combine($false)
if ($hTrue -eq $hFalse) { Write-Host "FAIL: HashCode true vs false failed"; exit 1 }
Write-Host "PASS"; exit 0
