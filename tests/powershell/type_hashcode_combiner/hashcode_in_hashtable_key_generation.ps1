# vybe-test: powershell/type_hashcode_combiner/hashcode_in_hashtable_key_generation
$h = @{}
$key1 = [System.HashCode]::Combine(10, 20)
$h[$key1] = "Stored"
if ($h[$key1] -ne "Stored") { Write-Host "FAIL: HashCode as hashtable key failed"; exit 1 }
Write-Host "PASS"; exit 0
