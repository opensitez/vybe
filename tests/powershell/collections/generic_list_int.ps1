# vybe-test: powershell/collections/generic_list_int
$list = [System.Collections.Generic.List[int]]::new()
$list.Add(10)
$list.Add(20)
$list.Add(30)
if ($list.Count -ne 3) { Write-Host "FAIL: count"; exit 1 }
$list.Remove(20)
if ($list.Count -ne 2) { Write-Host "FAIL: after remove"; exit 1 }
if ($list[0] -ne 10) { Write-Host "FAIL: [0]"; exit 1 }
if ($list[1] -ne 30) { Write-Host "FAIL: [1]"; exit 1 }
Write-Host "PASS"
exit 0
