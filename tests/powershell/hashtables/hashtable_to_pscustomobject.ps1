# vybe-test: powershell/hashtables/hashtable_to_pscustomobject
$h = @{ FirstName = "John"; LastName = "Doe"; Age = 40 }
$obj = [PSCustomObject]$h
if ($obj.FirstName -ne "John") { Write-Host "FAIL: FirstName"; exit 1 }
if ($obj.LastName  -ne "Doe")  { Write-Host "FAIL: LastName";  exit 1 }
if ($obj.Age       -ne 40)     { Write-Host "FAIL: Age";       exit 1 }
Write-Host "PASS"
exit 0
