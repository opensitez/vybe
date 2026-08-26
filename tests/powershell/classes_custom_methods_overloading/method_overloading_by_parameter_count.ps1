# vybe-test: powershell/classes_custom_methods_overloading/method_overloading_by_parameter_count
class Multiplier {
    [int]Multiply([int]$a) { return $a * 2 }
    [int]Multiply([int]$a, [int]$b) { return $a * $b }
    [int]Multiply([int]$a, [int]$b, [int]$c) { return $a * $b * $c }
}
$m = [Multiplier]::new()
$r1 = $m.Multiply(5)
$r2 = $m.Multiply(5, 3)
$r3 = $m.Multiply(5, 3, 2)
if ($r1 -ne 10 -or $r2 -ne 15 -or $r3 -ne 30) {
    Write-Host "FAIL: Method overloading by parameter count failed"
    exit 1
}
Write-Host "PASS"
exit 0
