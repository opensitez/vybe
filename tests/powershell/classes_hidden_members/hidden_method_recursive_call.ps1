# vybe-test: powershell/classes_hidden_members/hidden_method_recursive_call
class FactorialCalc {
    [int]Compute([int]$n) { return $this.Fact($n) }
    hidden [int]Fact([int]$n) {
        if ($n -le 1) { return 1 }
        return $n * $this.Fact($n - 1)
    }
}
$fc = [FactorialCalc]::new()
$res = $fc.Compute(5)
if ($res -ne 120) {
    Write-Host "FAIL: Hidden recursive method failed, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
