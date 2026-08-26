# vybe-test: powershell/classes_generic_method_invocations/generic_method_invocation_11
$list = [System.Collections.Generic.List[int]]::new()
$list.Add(11)
$list.Add(22)
$arr = $list.ToArray()
if ($arr.Length -ne 2 -or $arr[0] -ne 11 -or $arr[1] -ne 22) { Write-Host "FAIL: Generic method invocation failed"; exit 1 }
Write-Host "PASS"; exit 0
