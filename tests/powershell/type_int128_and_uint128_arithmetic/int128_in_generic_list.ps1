# vybe-test: powershell/type_int128_and_uint128_arithmetic/int128_in_generic_list
$list = [System.Collections.Generic.List[System.Int128]]::new()
$list.Add([System.Int128]::Parse("100"))
$list.Add([System.Int128]::Parse("200"))
if ($list.Count -ne 2 -or $list[1].ToString() -ne "200") { Write-Host "FAIL: Int128 in List failed"; exit 1 }
Write-Host "PASS"; exit 0
