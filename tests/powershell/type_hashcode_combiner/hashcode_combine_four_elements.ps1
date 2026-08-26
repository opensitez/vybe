# vybe-test: powershell/type_hashcode_combiner/hashcode_combine_four_elements
$h1 = [System.HashCode]::Combine("a", "b", "c", "d")
$h2 = [System.HashCode]::Combine("a", "b", "c", "d")
if ($h1 -ne $h2) { Write-Host "FAIL: HashCode.Combine 4 failed"; exit 1 }
Write-Host "PASS"; exit 0
