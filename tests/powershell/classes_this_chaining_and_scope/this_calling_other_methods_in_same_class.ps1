# vybe-test: powershell/classes_this_chaining_and_scope/this_calling_other_methods_in_same_class
class MathEngine {
    [int]Square([int]$x) { return $x * $x }
    [int]SumOfSquares([int]$a, [int]$b) {
        return $this.Square($a) + $this.Square($b)
    }
}
$me = [MathEngine]::new()
$res = $me.SumOfSquares(3, 4) # 9 + 16 = 25
if ($res -ne 25) {
    Write-Host "FAIL: `$this method call failed, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
