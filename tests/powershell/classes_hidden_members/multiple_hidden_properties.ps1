# vybe-test: powershell/classes_hidden_members/multiple_hidden_properties
class MultiHidden {
    hidden [int]$A = 1
    hidden [int]$B = 2
    [int]$C = 3
    [int]SumAll() { return $this.A + $this.B + $this.C }
}
$m = [MultiHidden]::new()
if ($m.SumAll() -ne 6) {
    Write-Host "FAIL: Multiple hidden properties sum failed"
    exit 1
}
Write-Host "PASS"
exit 0
