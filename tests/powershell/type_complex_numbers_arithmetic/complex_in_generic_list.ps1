# vybe-test: powershell/type_complex_numbers_arithmetic/complex_in_generic_list
$list = [System.Collections.Generic.List[System.Numerics.Complex]]::new()
$list.Add([System.Numerics.Complex]::new(10.0, 20.0))
if ($list.Count -ne 1 -or $list[0].Real -ne 10.0) { Write-Host "FAIL: Complex in List failed"; exit 1 }
Write-Host "PASS"; exit 0
