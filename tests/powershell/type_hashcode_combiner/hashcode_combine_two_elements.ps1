# vybe-test: powershell/type_hashcode_combiner/hashcode_combine_two_elements
$h = [System.HashCode]::Combine("test", 42)
if ($h -eq 0 -and $h -ne [System.HashCode]::Combine("test", 42)) { Write-Host "FAIL: HashCode.Combine 2 failed"; exit 1 }
Write-Host "PASS"; exit 0
