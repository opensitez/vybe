# vybe-test: powershell/type_hashcode_combiner/hashcode_combine_seven_elements
$h1 = [System.HashCode]::Combine(1, 2, 3, 4, 5, 6, 7)
$h2 = [System.HashCode]::Combine(1, 2, 3, 4, 5, 6, 7)
if ($h1 -ne $h2) { Write-Host "FAIL: HashCode.Combine 7 failed"; exit 1 }
Write-Host "PASS"; exit 0
