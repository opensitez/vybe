# vybe-test: powershell/classes_generic_method_invocations/generic_method_invocation_17
$list = [System.Collections.Generic.List[int]]::new()
$list.Add(17)
$list.Add(34)
$arr = $list.ToArray()
if ($arr.Length -ne 2 -or $arr[0] -ne 17 -or $arr[1] -ne 34) { Write-Host "FAIL: Generic method invocation failed"; exit 1 }
Write-Host "PASS"; exit 0
