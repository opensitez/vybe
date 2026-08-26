# vybe-test: powershell/classes_constructor_overloading/multiple_instances_independent_state
class Counter {
    [int]$Count
    Counter([int]$c) { $this.Count = $c }
}
$c1 = [Counter]::new(1)
$c2 = [Counter]::new(2)
if ($c1.Count -ne 1 -or $c2.Count -ne 2) {
    Write-Host "FAIL: Independent constructor state failed"
    exit 1
}
Write-Host "PASS"
exit 0
