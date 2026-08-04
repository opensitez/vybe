# vybe-test: powershell/classes/class_method
class Calculator {
    [int]Add([int]$a, [int]$b) { return $a + $b }
    [int]Multiply([int]$a, [int]$b) { return $a * $b }
}
$calc = [Calculator]::new()
$sum = $calc.Add(6, 7)
$prod = $calc.Multiply(4, 5)
if ($sum -ne 13) { Write-Host "FAIL: Add expected 13, got $sum"; exit 1 }
if ($prod -ne 20) { Write-Host "FAIL: Multiply expected 20, got $prod"; exit 1 }
Write-Host "PASS"
exit 0
