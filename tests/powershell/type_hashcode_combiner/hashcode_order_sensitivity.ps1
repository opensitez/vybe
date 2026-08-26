# vybe-test: powershell/type_hashcode_combiner/hashcode_order_sensitivity
$h1 = [System.HashCode]::Combine("first", "second")
$h2 = [System.HashCode]::Combine("second", "first")
if ($h1 -eq $h2) { Write-Host "FAIL: HashCode should be sensitive to argument order"; exit 1 }
Write-Host "PASS"; exit 0
