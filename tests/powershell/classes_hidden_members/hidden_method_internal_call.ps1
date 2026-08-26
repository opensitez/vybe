# vybe-test: powershell/classes_hidden_members/hidden_method_internal_call
class Calculator {
    [int]Compute([int]$x) {
        return $this.InternalHelper($x) * 2
    }
    hidden [int]InternalHelper([int]$x) {
        return $x + 10
    }
}
$c = [Calculator]::new()
$res = $c.Compute(5) # (5 + 10) * 2 = 30
if ($res -ne 30) {
    Write-Host "FAIL: Hidden method call from public method failed, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
