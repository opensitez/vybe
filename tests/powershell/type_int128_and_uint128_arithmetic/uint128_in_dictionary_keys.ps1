# vybe-test: powershell/type_int128_and_uint128_arithmetic/uint128_in_dictionary_keys
$d = [System.Collections.Generic.Dictionary[System.UInt128, string]]::new()
$key = [System.UInt128]::Parse("100000000000000000000")
$d.Add($key, "LargeKey")
if ($d[$key] -ne "LargeKey") { Write-Host "FAIL: UInt128 as Dictionary key failed"; exit 1 }
Write-Host "PASS"; exit 0
