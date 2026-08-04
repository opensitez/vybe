# vybe-test: powershell/objects/object_get_type
$n   = 42
$s   = "hello"
$arr = @(1,2,3)
if ($n.GetType().Name   -ne "Int32")    { Write-Host "FAIL: int type";   exit 1 }
if ($s.GetType().Name   -ne "String")   { Write-Host "FAIL: string type"; exit 1 }
if ($arr.GetType().Name -ne "Object[]") { Write-Host "FAIL: array type"; exit 1 }
Write-Host "PASS"
exit 0
