# vybe-test: powershell/classes_this_chaining_and_scope/this_state_preservation_across_calls
class Accumulator {
    [int]$Sum = 0
    [int]Accumulate([int]$v) {
        $this.Sum += $v
        return $this.Sum
    }
}
$acc = [Accumulator]::new()
$acc.Accumulate(10)
$acc.Accumulate(20)
$acc.Accumulate(30)
if ($acc.Sum -ne 60) {
    Write-Host "FAIL: Accumulator state preservation failed, got $($acc.Sum)"
    exit 1
}
Write-Host "PASS"
exit 0
