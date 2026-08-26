# vybe-test: powershell/type_hashcode_combiner/hashcode_with_null_arguments
$h1 = [System.HashCode]::Combine([string]$null, "valid")
$h2 = [System.HashCode]::Combine([string]$null, "valid")
if ($h1 -ne $h2) { Write-Host "FAIL: HashCode with null arguments failed"; exit 1 }
Write-Host "PASS"; exit 0
