# vybe-test: powershell/classes_generic_method_invocations/generic_method_invocation_2
$list = [System.Collections.Generic.List[int]]::new()
$list.Add(2)
$list.Add(4)
$arr = $list.ToArray()
if ($arr.Length -ne 2 -or $arr[0] -ne 2 -or $arr[1] -ne 4) { Write-Host "FAIL: Generic method invocation failed"; exit 1 }
Write-Host "PASS"; exit 0
