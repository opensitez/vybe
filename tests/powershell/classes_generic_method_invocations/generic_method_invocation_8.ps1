# vybe-test: powershell/classes_generic_method_invocations/generic_method_invocation_8
$list = [System.Collections.Generic.List[int]]::new()
$list.Add(8)
$list.Add(16)
$arr = $list.ToArray()
if ($arr.Length -ne 2 -or $arr[0] -ne 8 -or $arr[1] -ne 16) { Write-Host "FAIL: Generic method invocation failed"; exit 1 }
Write-Host "PASS"; exit 0
