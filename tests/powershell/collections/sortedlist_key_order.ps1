# vybe-test: powershell/collections/sortedlist
$sl = [System.Collections.Generic.SortedList[int,string]]::new()
$sl.Add(3, "three")
$sl.Add(1, "one")
$sl.Add(2, "two")
# SortedList iterates in key order
$keys = $sl.Keys
if ($keys[0] -ne 1) { Write-Host "FAIL: first key should be 1"; exit 1 }
if ($keys[1] -ne 2) { Write-Host "FAIL: second key should be 2"; exit 1 }
if ($sl[2] -ne "two") { Write-Host "FAIL: value lookup"; exit 1 }
Write-Host "PASS"
exit 0
