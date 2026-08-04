# vybe-test: powershell/comparison/contains_in_operator
$arr = @("apple", "banana", "cherry")
if ($arr -notcontains "banana") { Write-Host "FAIL: should contain banana"; exit 1 }
if ($arr -contains "grape")     { Write-Host "FAIL: should not contain grape"; exit 1 }
if ("banana" -notin $arr) { Write-Host "FAIL: -notin banana"; exit 1 }
if ("grape"  -in $arr)    { Write-Host "FAIL: -in grape";     exit 1 }
Write-Host "PASS"
exit 0
