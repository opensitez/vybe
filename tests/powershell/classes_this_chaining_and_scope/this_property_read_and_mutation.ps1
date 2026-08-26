# vybe-test: powershell/classes_this_chaining_and_scope/this_property_read_and_mutation
class CounterClass {
    [int]$Count = 0
    [void]Increment() { $this.Count++ }
    [void]Add([int]$amount) { $this.Count += $amount }
}
$c = [CounterClass]::new()
$c.Increment()
$c.Add(10)
if ($c.Count -ne 11) {
    Write-Host "FAIL: `$this property mutation failed, got $($c.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
