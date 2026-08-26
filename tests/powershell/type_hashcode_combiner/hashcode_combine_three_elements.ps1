# vybe-test: powershell/type_hashcode_combiner/hashcode_combine_three_elements
$h1 = [System.HashCode]::Combine(1, 2, 3)
$h2 = [System.HashCode]::Combine(1, 2, 3)
$h3 = [System.HashCode]::Combine(1, 2, 4)
if ($h1 -ne $h2 -or $h1 -eq $h3) { Write-Host "FAIL: HashCode.Combine 3 consistency failed"; exit 1 }
Write-Host "PASS"; exit 0
