# vybe-test: powershell/classes_this_chaining_and_scope/this_in_nested_loop_inside_method
class MatrixHelper {
    [int]$Scale = 2
    [int]SumScaled([int[]]$a, [int[]]$b) {
        $total = 0
        foreach ($x in $a) {
            foreach ($y in $b) {
                $total += ($x + $y) * $this.Scale
            }
        }
        return $total
    }
}
$mh = [MatrixHelper]::new()
$res = $mh.SumScaled(@(1, 2), @(3, 4)) # (4+5+5+6)*2 = 20*2 = 40
if ($res -ne 40) {
    Write-Host "FAIL: `$this in nested loop failed, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
