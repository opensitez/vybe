# vybe-test: powershell/operators/type_cast_operator
$str  = [string]42
$int  = [int]"7"
$bool = [bool]0
$bool2= [bool]1
if ($str  -ne "42")   { Write-Host "FAIL: int->string"; exit 1 }
if ($int  -ne 7)      { Write-Host "FAIL: string->int"; exit 1 }
if ($bool  -ne $false){ Write-Host "FAIL: 0->bool";     exit 1 }
if ($bool2 -ne $true) { Write-Host "FAIL: 1->bool";     exit 1 }
Write-Host "PASS"
exit 0
