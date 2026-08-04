# vybe-test: powershell/functions/function_recursive_closures
function Make-Adder([int]$n) {
    return { param($x) $x + $n }.GetNewClosure()
}
$add5  = Make-Adder 5
$add10 = Make-Adder 10
if ((& $add5 3)  -ne 8)  { Write-Host "FAIL: add5(3)";  exit 1 }
if ((& $add10 3) -ne 13) { Write-Host "FAIL: add10(3)"; exit 1 }
if ((& $add5 0)  -ne 5)  { Write-Host "FAIL: add5(0)";  exit 1 }
Write-Host "PASS"
exit 0
