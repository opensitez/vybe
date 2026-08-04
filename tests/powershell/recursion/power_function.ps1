# vybe-test: powershell/recursion/power_function
function Pow([double]$base, [int]$exp) {
    if ($exp -eq 0) { return 1 }
    if ($exp -lt 0) { return 1 / (Pow $base (-$exp)) }
    return $base * (Pow $base ($exp - 1))
}
if ((Pow 2 10) -ne 1024)     { Write-Host "FAIL: 2^10"; exit 1 }
if ((Pow 3 4)  -ne 81)       { Write-Host "FAIL: 3^4";  exit 1 }
if ((Pow 5 0)  -ne 1)        { Write-Host "FAIL: x^0";  exit 1 }
Write-Host "PASS"
exit 0
